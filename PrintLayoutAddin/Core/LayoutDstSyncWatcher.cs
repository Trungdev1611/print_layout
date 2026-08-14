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
    /// When the user deletes a layout that PLAYOUT created (see <see cref="PlayoutLayoutStamp"/>),
    /// offer to remove the matching sheet(s) from a Sheet Set they pick. Layouts without the
    /// stamp — hand-made tabs, layouts from other drawings — are deleted silently.
    ///
    /// Anti-loop rules, each one earned:
    ///  * exactly one path reaches the prompt: erase -&gt; Idle -&gt; <see cref="SyncCommand"/> -&gt; prompt.
    ///    CommandEnded is deliberately NOT hooked: it fires for the sync command itself and
    ///    re-entered the prompt;
    ///  * CommandWillStart ignores the sync command for the same reason;
    ///  * a batch is closed for good in finally — Pending is cleared and Idle is NOT re-hooked,
    ///    so anything raised while the dialog was open (Undo/unerase, the second erase event)
    ///    cannot start another round;
    ///  * <see cref="OnObjectErased"/> checks <c>e.Erased</c>, so Undo (an unerase) does not
    ///    look like a fresh deletion.
    /// </summary>
    public static class LayoutDstSyncWatcher
    {
        public const string SyncCommand = "PLDSTLAYOUTSYNC";

        /// <summary>A name re-queued within this window is treated as an echo of the batch just shown.</summary>
        private static readonly TimeSpan EchoWindow = TimeSpan.FromSeconds(20);

        private static int _suppress;
        private static bool _started;
        private static bool _idleHooked;
        private static bool _commandScheduled;
        private static bool _processing;
        private static readonly object Gate = new object();

        private static readonly List<string> Pending = new List<string>();
        private static string _pendingDstHint = "";
        private static Document _pendingDoc;

        private static readonly Dictionary<string, DateTime> RecentlyPrompted =
            new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Stamped layout name -&gt; DST path from its stamp, for the active document.
        /// LayoutRemoved only hands us a name, and by then the Layout is gone — so the
        /// stamps are read up front, before the command that may delete them runs.
        /// </summary>
        private static readonly Dictionary<string, string> Stamped =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static Document _snapshotDoc;

        private static LayoutManager _layoutManager;
        private static bool _docsHooked;
        private static readonly HashSet<Document> HookedDocs = new HashSet<Document>();

        public static void Start()
        {
            if (_started) return;
            try
            {
                if (!_docsHooked)
                {
                    AcadApp.DocumentManager.DocumentActivated += OnDocumentActivated;
                    AcadApp.DocumentManager.DocumentCreated += OnDocumentCreated;
                    AcadApp.DocumentManager.DocumentDestroyed += OnDocumentDestroyed;
                    _docsHooked = true;
                }
                foreach (Document d in AcadApp.DocumentManager)
                    HookDocument(d);
                HookLayoutManager();
                RefreshSnapshot();
                _started = true;
            }
            catch { }
        }

        public static void Stop()
        {
            if (!_started) return;
            UnhookLayoutManager();
            if (_docsHooked)
            {
                try { AcadApp.DocumentManager.DocumentActivated -= OnDocumentActivated; } catch { }
                try { AcadApp.DocumentManager.DocumentCreated -= OnDocumentCreated; } catch { }
                try { AcadApp.DocumentManager.DocumentDestroyed -= OnDocumentDestroyed; } catch { }
                _docsHooked = false;
            }
            lock (Gate)
            {
                foreach (var doc in HookedDocs.ToList())
                    UnhookDocument(doc);
                HookedDocs.Clear();
            }
            UnhookIdle();
            lock (Gate)
            {
                ClearBatch();
                Stamped.Clear();
                _snapshotDoc = null;
                RecentlyPrompted.Clear();
                _commandScheduled = false;
                _processing = false;
            }
            _started = false;
        }

        // ---------------------------------------------------------------- suppression

        /// <summary>Deletions inside the scope are ours (PLAYOUT overwrite, frame cleanup) — never prompt.</summary>
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

        // ---------------------------------------------------------------- documents

        private static void OnDocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            HookDocument(e?.Document);
            HookLayoutManager();
            RefreshSnapshot();
        }

        private static void OnDocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            HookDocument(e?.Document);
        }

        private static void OnDocumentDestroyed(object sender, EventArgs e)
        {
            lock (Gate)
            {
                var gone = HookedDocs.Where(d =>
                {
                    try { return d == null || d.IsDisposed; }
                    catch { return true; }
                }).ToList();
                foreach (var d in gone)
                    HookedDocs.Remove(d);

                if (_pendingDoc != null && !HookedDocs.Contains(_pendingDoc))
                    ClearBatch();
            }
        }

        private static void HookDocument(Document doc)
        {
            if (doc == null) return;
            lock (Gate)
            {
                if (HookedDocs.Contains(doc)) return;
                HookedDocs.Add(doc);
            }
            try
            {
                doc.CommandWillStart += OnCommandWillStart;
                doc.Database.ObjectErased += OnObjectErased;
            }
            catch { }
        }

        private static void UnhookDocument(Document doc)
        {
            if (doc == null) return;
            try { doc.CommandWillStart -= OnCommandWillStart; } catch { }
            try { doc.Database.ObjectErased -= OnObjectErased; } catch { }
        }

        // ---------------------------------------------------------------- stamp snapshot

        /// <summary>Re-read the PLAYOUT stamps of the active drawing.</summary>
        public static void RefreshSnapshot()
        {
            try
            {
                var doc = AcadApp.DocumentManager?.MdiActiveDocument;
                var map = PlayoutLayoutStamp.Snapshot(doc?.Database);
                lock (Gate)
                {
                    _snapshotDoc = doc;
                    Stamped.Clear();
                    foreach (var kv in map)
                        Stamped[kv.Key] = kv.Value;
                }
            }
            catch { }
        }

        private static void OnCommandWillStart(object sender, CommandEventArgs e)
        {
            if (IsSyncCommand(e)) return;
            if (_processing) return;
            RefreshSnapshot();
        }

        private static bool IsSyncCommand(CommandEventArgs e)
        {
            string name = null;
            try { name = e?.GlobalCommandName; } catch { }
            return !string.IsNullOrWhiteSpace(name)
                && name.TrimStart('-', '_', '\'')
                       .Equals(SyncCommand, StringComparison.OrdinalIgnoreCase);
        }

        private static void HookLayoutManager()
        {
            try
            {
                UnhookLayoutManager();
                var lm = LayoutManager.Current;
                if (lm == null) return;
                lm.LayoutRemoved += OnLayoutRemoved;
                lm.LayoutCreated += OnLayoutCreated;
                lm.LayoutRenamed += OnLayoutRenamed;
                _layoutManager = lm;
            }
            catch { }
        }

        private static void UnhookLayoutManager()
        {
            if (_layoutManager == null) return;
            try { _layoutManager.LayoutRemoved -= OnLayoutRemoved; } catch { }
            try { _layoutManager.LayoutCreated -= OnLayoutCreated; } catch { }
            try { _layoutManager.LayoutRenamed -= OnLayoutRenamed; } catch { }
            _layoutManager = null;
        }

        private static void OnLayoutCreated(object sender, Autodesk.AutoCAD.DatabaseServices.LayoutEventArgs e)
        {
            // A copy of a stamped layout carries the stamp; a brand-new blank tab does not.
            // Either way the snapshot has to learn about the new name before it can be deleted.
            if (_processing) return;
            AddToSnapshotIfStamped(e?.Name);
        }

        private static void OnLayoutRenamed(object sender, LayoutRenamedEventArgs e)
        {
            if (_processing) return;
            lock (Gate)
            {
                // LayoutRenamedEventArgs.Name is the old name; NewName is the new one.
                if (!string.IsNullOrWhiteSpace(e?.Name))
                    Stamped.Remove(e.Name.Trim());
            }
            AddToSnapshotIfStamped(e?.NewName);
        }

        private static void AddToSnapshotIfStamped(string layoutName)
        {
            if (string.IsNullOrWhiteSpace(layoutName)) return;
            layoutName = layoutName.Trim();
            try
            {
                var doc = AcadApp.DocumentManager?.MdiActiveDocument;
                if (doc?.Database == null) return;
                if (!PlayoutLayoutStamp.TryReadByName(doc.Database, layoutName, out var dst)) return;

                lock (Gate)
                {
                    _snapshotDoc = doc;
                    Stamped[layoutName] = dst ?? "";
                }
            }
            catch { }
        }

        // ---------------------------------------------------------------- deletion capture

        private static void OnObjectErased(object sender, ObjectErasedEventArgs e)
        {
            if (_suppress > 0 || _processing) return;
            if (e == null || !e.Erased) return;   // unerase (Undo) is not a deletion

            var layout = e.DBObject as Layout;
            if (layout == null) return;

            string name = null;
            try { name = layout.LayoutName; } catch { }
            if (string.IsNullOrWhiteSpace(name)) return;

            // Authoritative: the stamp is still readable on the object being erased.
            // Fall back to the pre-command snapshot rather than assuming "not ours".
            if (!PlayoutLayoutStamp.TryReadFromErased(layout, out var dst)
                && !TryGetSnapshotStamp(name, out dst))
                return;

            Queue(name, dst);
        }

        private static void OnLayoutRemoved(object sender, Autodesk.AutoCAD.DatabaseServices.LayoutEventArgs e)
        {
            if (_suppress > 0 || _processing) return;
            // The Layout object is already gone here, so the snapshot is all we have.
            // Not stamped => not ours => delete silently, no prompt.
            if (!TryGetSnapshotStamp(e?.Name, out var dst)) return;
            Queue(e.Name, dst);
        }

        private static bool TryGetSnapshotStamp(string layoutName, out string dstPath)
        {
            dstPath = "";
            if (string.IsNullOrWhiteSpace(layoutName)) return false;

            Document active = null;
            try { active = AcadApp.DocumentManager?.MdiActiveDocument; } catch { }

            lock (Gate)
            {
                // A snapshot taken in another drawing says nothing about this one.
                if (!ReferenceEquals(_snapshotDoc, active)) return false;
                return Stamped.TryGetValue(layoutName.Trim(), out dstPath);
            }
        }

        private static void Queue(string layoutName, string dstHint)
        {
            if (_suppress > 0 || _processing) return;

            layoutName = (layoutName ?? "").Trim();
            if (string.IsNullOrWhiteSpace(layoutName)) return;
            if (string.Equals(layoutName, "Model", StringComparison.OrdinalIgnoreCase)) return;

            var doc = AcadApp.DocumentManager?.MdiActiveDocument;

            lock (Gate)
            {
                if (WasRecentlyPrompted(layoutName)) return;

                // A batch belongs to one drawing; switching drawings starts a new one.
                if (_pendingDoc != null && !ReferenceEquals(_pendingDoc, doc))
                    ClearBatch();
                _pendingDoc = doc;

                Stamped.Remove(layoutName);

                if (!Pending.Contains(layoutName, StringComparer.OrdinalIgnoreCase))
                    Pending.Add(layoutName);

                if (string.IsNullOrWhiteSpace(_pendingDstHint) && !string.IsNullOrWhiteSpace(dstHint))
                    _pendingDstHint = dstHint;
            }

            HookIdle();
        }

        private static void ClearBatch()
        {
            Pending.Clear();
            _pendingDstHint = "";
            _pendingDoc = null;
        }

        private static bool WasRecentlyPrompted(string layoutName)
        {
            if (!RecentlyPrompted.TryGetValue(layoutName, out var when)) return false;
            if (DateTime.UtcNow - when < EchoWindow) return true;
            RecentlyPrompted.Remove(layoutName);
            return false;
        }

        private static void MarkPrompted(IEnumerable<string> layoutNames)
        {
            var now = DateTime.UtcNow;
            foreach (var name in layoutNames)
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                RecentlyPrompted[name.Trim()] = now;
            }

            var stale = RecentlyPrompted
                .Where(kv => now - kv.Value > TimeSpan.FromMinutes(10))
                .Select(kv => kv.Key)
                .ToList();
            foreach (var k in stale) RecentlyPrompted.Remove(k);
        }

        // ---------------------------------------------------------------- idle -> command

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

        private static void OnIdle(object sender, EventArgs e)
        {
            UnhookIdle();

            bool schedule;
            lock (Gate)
            {
                schedule = Pending.Count > 0 && !_commandScheduled && !_processing;
                if (schedule) _commandScheduled = true;
            }
            if (!schedule) return;

            try
            {
                Document doc;
                lock (Gate) doc = _pendingDoc ?? AcadApp.DocumentManager?.MdiActiveDocument;
                if (doc == null)
                {
                    lock (Gate)
                    {
                        _commandScheduled = false;
                        ClearBatch();
                    }
                    return;
                }
                // MessageBox inside Idle re-enters the message pump; run the prompt as a command.
                doc.SendStringToExecute(SyncCommand + " ", true, false, false);
            }
            catch
            {
                lock (Gate)
                {
                    _commandScheduled = false;
                    ClearBatch();
                }
            }
        }

        // ---------------------------------------------------------------- the prompt

        /// <summary>Entry point of <see cref="SyncCommand"/> — the only place that shows UI.</summary>
        public static void ProcessPendingFromCommand()
        {
            List<string> layouts;
            string dstHint;
            Document doc;

            lock (Gate)
            {
                _commandScheduled = false;
                if (_processing) return;

                layouts = Pending
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                dstHint = _pendingDstHint;
                doc = _pendingDoc;
                ClearBatch();

                if (layouts.Count == 0) return;
                MarkPrompted(layouts);
                _processing = true;
            }

            try
            {
                Prompt(layouts, dstHint, doc);
            }
            finally
            {
                // Close the batch for good. Events raised while the dialog was up (the second
                // erase event, an Undo) are dropped here and Idle is NOT re-hooked — that pair
                // is what used to spin the dialog forever.
                lock (Gate)
                {
                    _processing = false;
                    _commandScheduled = false;
                    ClearBatch();
                }
                RefreshSnapshot();
            }
        }

        private static void Prompt(List<string> layouts, string dstHint, Document doc)
        {
            string list = layouts.Count <= 8
                ? string.Join(", ", layouts)
                : string.Join(", ", layouts.Take(8)) + ", …";

            var owner = AcadOwner.Get();
            var pick = MessageBox.Show(
                owner,
                "Deleted layout(s) created by Build Layouts:\n\n"
                + list
                + "\n\nAlso remove the matching sheet(s) from the linked Sheet Set (.dst)?\n\n"
                + "Yes — pick the .dst file\n"
                + "No — leave the Sheet Set unchanged",
                "Print Layout — update Sheet Set?",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (pick != DialogResult.Yes) return;

            string dstPath = PromptDstFile(dstHint, doc, owner);
            if (string.IsNullOrWhiteSpace(dstPath)) return;

            try
            {
                if (!SheetSetService.TryRemoveSheetsByLayoutNames(dstPath, layouts, out var message))
                {
                    MessageBox.Show(
                        owner,
                        message,
                        "Print Layout — Sheet Set update",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                var note = (message ?? "Sheet(s) removed.")
                    + "\n\nLayout(s): " + list
                    + "\n\nDST file:\n" + dstPath
                    + "\n\nRefresh Sheet Set Manager if the tree still shows the old sheets.";
                try { doc?.Editor?.WriteMessage("\n" + note.Replace("\n\n", " | ")); } catch { }
                MessageBox.Show(
                    owner,
                    note,
                    "Print Layout — sheet removed from Sheet Set",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    owner,
                    ex.Message,
                    "Print Layout — Sheet Set update",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static string PromptDstFile(string dstHint, Document doc, IWin32Window owner)
        {
            string dwgPath = null;
            try { dwgPath = doc?.Name; } catch { }

            string initialFile = FirstExistingDst(dstHint, SafeDefaultDst(dwgPath));
            string initialDir = "";
            try
            {
                initialDir = !string.IsNullOrWhiteSpace(initialFile)
                    ? Path.GetDirectoryName(initialFile)
                    : (File.Exists(dwgPath) ? Path.GetDirectoryName(dwgPath) : "");
            }
            catch { }

            using (var ofd = new OpenFileDialog
            {
                Title = "Choose the Sheet Set (.dst) to update",
                Filter = "Sheet Set (*.dst)|*.dst",
                DefaultExt = "dst",
                CheckFileExists = true,
                FileName = initialFile ?? "",
                InitialDirectory = Directory.Exists(initialDir) ? initialDir : "",
            })
            {
                if (ofd.ShowDialog(owner) != DialogResult.OK) return "";

                var path = ofd.FileName?.Trim() ?? "";
                if (!path.EndsWith(".dst", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(
                        owner,
                        "That is not a .dst file.",
                        "Print Layout — Sheet Set update",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return "";
                }
                if (SheetSetService.TryRead(path) == null)
                {
                    MessageBox.Show(
                        owner,
                        "Could not read that file as a Sheet Set.\n\n" + path,
                        "Print Layout — Sheet Set update",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return "";
                }
                return path;
            }
        }

        private static string SafeDefaultDst(string dwgPath)
        {
            try
            {
                return string.IsNullOrWhiteSpace(dwgPath)
                    ? null
                    : PublishPaths.DefaultDstPath(dwgPath);
            }
            catch { return null; }
        }

        private static string FirstExistingDst(params string[] candidates)
        {
            foreach (var c in candidates)
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(c) && File.Exists(c)) return c;
                }
                catch { }
            }
            return null;
        }

        /// <summary>Parents the WinForms dialogs to AutoCAD so they cannot slip behind it.</summary>
        private sealed class AcadOwner : IWin32Window
        {
            private AcadOwner(IntPtr handle) { Handle = handle; }

            public IntPtr Handle { get; }

            public static IWin32Window Get()
            {
                try
                {
                    var handle = AcadApp.MainWindow?.Handle ?? IntPtr.Zero;
                    if (handle != IntPtr.Zero) return new AcadOwner(handle);
                }
                catch { }
                return null;
            }
        }
    }
}
