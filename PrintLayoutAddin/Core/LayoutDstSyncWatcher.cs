using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace PrintLayoutAddin.Core
{
    /// <summary>
    /// When the user deletes a layout, offer to remove matching sheet(s) from the DST.
    /// Prompt runs via command <c>PLDSTLAYOUTSYNC</c> (not MessageBox inside Idle)
    /// to avoid AutoCAD Idle re-entrancy loops.
    /// </summary>
    public static class LayoutDstSyncWatcher
    {
        public const string SyncCommand = "PLDSTLAYOUTSYNC";

        private static int _suppress;
        private static bool _started;
        private static bool _idleHooked;
        private static bool _commandScheduled;
        private static bool _processing;
        private static readonly object Gate = new object();
        private static readonly List<PendingRemoval> Pending = new List<PendingRemoval>();
        private static readonly Dictionary<string, DateTime> RecentlyHandled =
            new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);

        private sealed class PendingRemoval
        {
            public string DstPath;
            public string LayoutName;
        }

        public static void Start()
        {
            if (_started) return;
            try
            {
                LayoutManager.Current.LayoutRemoved += OnLayoutRemoved;
                _started = true;
            }
            catch { }
        }

        public static void Stop()
        {
            if (!_started) return;
            try { LayoutManager.Current.LayoutRemoved -= OnLayoutRemoved; } catch { }
            UnhookIdle();
            lock (Gate)
            {
                Pending.Clear();
                RecentlyHandled.Clear();
                _commandScheduled = false;
                _processing = false;
            }
            _started = false;
        }

        public static IDisposable Suppress()
        {
            System.Threading.Interlocked.Increment(ref _suppress);
            return new SuppressScope();
        }

        private sealed class SuppressScope : IDisposable
        {
            private bool _done;
            public void Dispose()
            {
                if (_done) return;
                _done = true;
                System.Threading.Interlocked.Decrement(ref _suppress);
            }
        }

        private static string HandledKey(string dstPath, string layoutName)
        {
            return (dstPath ?? "").Trim() + "|" + (layoutName ?? "").Trim();
        }

        private static bool WasRecentlyHandled(string dstPath, string layoutName)
        {
            string key = HandledKey(dstPath, layoutName);
            lock (Gate)
            {
                if (!RecentlyHandled.TryGetValue(key, out var when)) return false;
                if ((DateTime.UtcNow - when).TotalSeconds < 60) return true;
                RecentlyHandled.Remove(key);
                return false;
            }
        }

        private static void MarkHandled(string dstPath, IEnumerable<string> layoutNames)
        {
            lock (Gate)
            {
                var now = DateTime.UtcNow;
                foreach (var name in layoutNames)
                {
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    RecentlyHandled[HandledKey(dstPath, name)] = now;
                }

                var stale = RecentlyHandled
                    .Where(kv => (now - kv.Value).TotalMinutes > 10)
                    .Select(kv => kv.Key)
                    .ToList();
                foreach (var k in stale) RecentlyHandled.Remove(k);
            }
        }

        private static void OnLayoutRemoved(object sender, Autodesk.AutoCAD.DatabaseServices.LayoutEventArgs e)
        {
            if (_suppress > 0 || _processing) return;
            if (e == null || string.IsNullOrWhiteSpace(e.Name)) return;

            string layoutName = e.Name.Trim();
            if (IsProtectedLayout(layoutName)) return;

            var doc = AcadApp.DocumentManager?.MdiActiveDocument;
            if (doc == null) return;

            string dwgPath = null;
            try { dwgPath = doc.Name; } catch { }
            if (string.IsNullOrWhiteSpace(dwgPath) || !File.Exists(dwgPath))
                return;

            string dstPath = PublishPaths.DefaultDstPath(dwgPath);
            if (!File.Exists(dstPath)) return;
            if (WasRecentlyHandled(dstPath, layoutName)) return;

            try
            {
                var existing = SheetSetService.TryRead(dstPath);
                if (existing == null) return;
                bool hit = existing.Sheets.Any(s =>
                    s != null
                    && (string.Equals(s.LayoutName, layoutName, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(s.Number, layoutName, StringComparison.OrdinalIgnoreCase)));
                if (!hit) return;
            }
            catch
            {
                return;
            }

            lock (Gate)
            {
                bool dup = Pending.Any(p =>
                    string.Equals(p.DstPath, dstPath, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(p.LayoutName, layoutName, StringComparison.OrdinalIgnoreCase));
                if (!dup)
                {
                    Pending.Add(new PendingRemoval
                    {
                        DstPath = dstPath,
                        LayoutName = layoutName,
                    });
                }
            }

            HookIdle();
        }

        private static bool IsProtectedLayout(string layoutName)
        {
            if (string.Equals(layoutName, "Model", StringComparison.OrdinalIgnoreCase))
                return true;
            string template = Config.Instance.TemplateLayout;
            if (!string.IsNullOrWhiteSpace(template)
                && string.Equals(layoutName, template, StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        private static void HookIdle()
        {
            if (_idleHooked) return;
            try
            {
                AcadApp.Idle += OnIdle;
                _idleHooked = true;
            }
            catch { }
        }

        private static void UnhookIdle()
        {
            if (!_idleHooked) return;
            try { AcadApp.Idle -= OnIdle; } catch { }
            _idleHooked = false;
        }

        /// <summary>
        /// Idle only schedules a command — never shows UI here (avoids Idle re-entrancy).
        /// </summary>
        private static void OnIdle(object sender, EventArgs e)
        {
            UnhookIdle();

            bool needSchedule;
            lock (Gate)
            {
                needSchedule = Pending.Count > 0 && !_commandScheduled && !_processing;
                if (needSchedule) _commandScheduled = true;
            }
            if (!needSchedule) return;

            try
            {
                var doc = AcadApp.DocumentManager?.MdiActiveDocument;
                if (doc == null)
                {
                    lock (Gate) _commandScheduled = false;
                    return;
                }
                doc.SendStringToExecute(SyncCommand + " ", true, false, false);
            }
            catch
            {
                lock (Gate) _commandScheduled = false;
            }
        }

        /// <summary>Called from <see cref="SyncCommand"/> — safe to show MessageBox.</summary>
        public static void ProcessPendingFromCommand()
        {
            List<PendingRemoval> batch;
            lock (Gate)
            {
                _commandScheduled = false;
                if (_processing)
                {
                    Pending.Clear();
                    return;
                }
                batch = Pending.ToList();
                Pending.Clear();
                if (batch.Count == 0) return;
                _processing = true;
            }

            try
            {
                // Open SSM once per successfully updated DST after the whole batch —
                // not after each layout.
                var reopenDsts = new List<string>();

                foreach (var group in batch.GroupBy(p => p.DstPath, StringComparer.OrdinalIgnoreCase))
                {
                    var layouts = group.Select(g => g.LayoutName)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    if (layouts.Count == 0) continue;

                    string dstPath = group.Key;

                    // Mark before UI so any stray events cannot re-queue this batch.
                    MarkHandled(dstPath, layouts);

                    string list = layouts.Count <= 8
                        ? string.Join(", ", layouts)
                        : string.Join(", ", layouts.Take(8)) + ", …";

                    var answer = MessageBox.Show(
                        "Layout(s) deleted:\n\n"
                        + list
                        + "\n\nRemove the matching sheet(s) from the Sheet Set (DST)?\n\n"
                        + dstPath,
                        "Print Layout — update DST?",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (answer != DialogResult.Yes) continue;

                    try
                    {
                        if (!SheetSetService.TryRemoveSheetsByLayoutNames(
                                dstPath, layouts, out var message))
                        {
                            MessageBox.Show(
                                message,
                                "Print Layout — DST update",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                        }
                        else
                        {
                            var doc = AcadApp.DocumentManager?.MdiActiveDocument;
                            doc?.Editor?.WriteMessage("\n" + message);
                            if (!string.IsNullOrWhiteSpace(dstPath)
                                && !reopenDsts.Any(p =>
                                    string.Equals(p, dstPath, StringComparison.OrdinalIgnoreCase)))
                                reopenDsts.Add(dstPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            ex.Message,
                            "Print Layout — DST update",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }

                // Previously this reopened each DST in Sheet Set Manager (ReloadForUser).
                // That close/release/reopen churn on a shared AcSm database is what left the
                // session unable to create a DST afterwards, so we only report the change —
                // SSM refreshes its own tree.
                foreach (var dstPath in reopenDsts)
                {
                    SheetSetAutoLog.Write(
                        AcadApp.DocumentManager?.MdiActiveDocument?.Editor,
                        AcadApp.DocumentManager?.MdiActiveDocument?.Name,
                        "after layout-delete: DST updated — " + dstPath
                        + " (refresh Sheet Set Manager to see it)");
                }
            }
            finally
            {
                lock (Gate)
                {
                    _processing = false;
                    // Drop anything re-queued for layouts we just handled.
                    Pending.RemoveAll(p => WasRecentlyHandled(p.DstPath, p.LayoutName));
                    if (Pending.Count > 0)
                        HookIdle();
                }
            }
        }
    }
}
