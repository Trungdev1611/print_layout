using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Autodesk.Windows;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace PrintLayoutAddin.UI
{
    public static class RibbonBuilder
    {
        private const string TabId = "PrintLayoutAddin_Tab";
        private const string PanelId = "PrintLayoutAddin_Panel";

        public static void Build()
        {
            var ribbon = ComponentManager.Ribbon;
            if (ribbon == null)
            {
                ComponentManager.ItemInitialized += OnItemInitialized;
                return;
            }
            BuildCore(ribbon);
        }

        private static void OnItemInitialized(object sender, RibbonItemEventArgs e)
        {
            if (ComponentManager.Ribbon == null) return;
            ComponentManager.ItemInitialized -= OnItemInitialized;
            BuildCore(ComponentManager.Ribbon);
        }

        private static void BuildCore(RibbonControl ribbon)
        {
            foreach (var t in ribbon.Tabs)
                if (t.Id == TabId) return;

            var tab = new RibbonTab
            {
                Title = "Print Layout",
                Id = TabId
            };

            var panelSource = new RibbonPanelSource { Title = "Layout Tools" };
            panelSource.Id = PanelId;

            var btnAuto = new RibbonButton
            {
                Text = "Auto\nFrames",
                ShowText = true,
                ShowImage = true,
                LargeImage = DrawAutoIcon(32),
                Image = DrawAutoIcon(16),
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                Description = "Detect nested frames inside an xref and generate matching native block references (PLAUTO).",
                CommandHandler = new RelayCommand(() => AcadApp.DocumentManager.MdiActiveDocument?
                    .SendStringToExecute("_PLAUTO ", true, false, true))
            };

            var btnStt = new RibbonButton
            {
                Text = "Number\nFrames",
                ShowText = true,
                ShowImage = true,
                LargeImage = DrawSttIcon(32),
                Image = DrawSttIcon(16),
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                Description = "Assign INNO-STT numbers to frames along a guide polyline (PLSTT).",
                CommandHandler = new RelayCommand(() => AcadApp.DocumentManager.MdiActiveDocument?
                    .SendStringToExecute("_PLSTT ", true, false, true))
            };

            var btnLayout = new RibbonButton
            {
                Text = "Build\nLayouts",
                ShowText = true,
                ShowImage = true,
                LargeImage = DrawLayoutIcon(32),
                Image = DrawLayoutIcon(16),
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                Description = "Generate layouts + viewport per frame (requires P1/P2 from Title Block Setup).",
                CommandHandler = new RelayCommand(() => AcadApp.DocumentManager.MdiActiveDocument?
                    .SendStringToExecute("_PLAYOUT ", true, false, true))
            };

            var btnPrint = new RibbonButton
            {
                Text = "Print\nPDF",
                ShowText = true,
                ShowImage = true,
                LargeImage = DrawPrintIcon(32),
                Image = DrawPrintIcon(16),
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                Description = "Print selected layouts or export them to one PDF/file (PLPRINT).",
                CommandHandler = new RelayCommand(() => AcadApp.DocumentManager.MdiActiveDocument?
                    .SendStringToExecute("_PLPRINT ", true, false, true))
            };

            var btnSheetSet = new RibbonButton
            {
                Text = "Create\nSheet Set",
                ShowText = true,
                ShowImage = true,
                LargeImage = DrawSheetSetIcon(32),
                Image = DrawSheetSetIcon(16),
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                Description = "Order the current DWG layouts, create/update a DST, or export PDF (PLSHEETSET).",
                CommandHandler = new RelayCommand(() => AcadApp.DocumentManager.MdiActiveDocument?
                    .SendStringToExecute("_PLSHEETSET ", true, false, true))
            };

            var btnFrameSetup = new RibbonButton
            {
                Text = "Title Block\nSetup",
                ShowText = true,
                ShowImage = true,
                LargeImage = DrawFrameSetupIcon(32),
                Image = DrawFrameSetupIcon(16),
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                Description = "Pick viewport P1/P2, typography, and run Auto Frame Setup (PLFRAME_SETUP).",
                CommandHandler = new RelayCommand(() => AcadApp.DocumentManager.MdiActiveDocument?
                    .SendStringToExecute("_PLFRAME_SETUP ", true, false, true))
            };

            var btnClean = new RibbonButton
            {
                Text = "Cleanup\nFrames",
                ShowText = true,
                ShowImage = true,
                LargeImage = DrawCleanIcon(32),
                Image = DrawCleanIcon(16),
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                Description = "Erase PLAUTO-generated native frame blocks and optionally their layouts (PLCLEAN).",
                CommandHandler = new RelayCommand(() => AcadApp.DocumentManager.MdiActiveDocument?
                    .SendStringToExecute("_PLCLEAN ", true, false, true))
            };

            var btnKeys = new RibbonButton
            {
                Text = "Keyboard\nShortcuts",
                ShowText = true,
                ShowImage = true,
                LargeImage = DrawKeysIcon(32),
                Image = DrawKeysIcon(16),
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                Description = "Assign custom command aliases, e.g. type 11 to run Print PDF (PLKEYS).",
                CommandHandler = new RelayCommand(() => AcadApp.DocumentManager.MdiActiveDocument?
                    .SendStringToExecute("_PLKEYS ", true, false, true))
            };

            var btnLicense = new RibbonButton
            {
                Text = "License",
                ShowText = true,
                ShowImage = true,
                LargeImage = DrawLicenseIcon(32),
                Image = DrawLicenseIcon(16),
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
                Description = "Activate or view the addin license (PLLICENSE).",
                CommandHandler = new RelayCommand(() => AcadApp.DocumentManager.MdiActiveDocument?
                    .SendStringToExecute("_PLLICENSE ", true, false, true))
            };

            panelSource.Items.Add(btnAuto);
            panelSource.Items.Add(btnFrameSetup);
            panelSource.Items.Add(btnStt);
            panelSource.Items.Add(btnLayout);
            panelSource.Items.Add(btnSheetSet);
            panelSource.Items.Add(btnPrint);
            panelSource.Items.Add(btnClean);
            panelSource.Items.Add(btnKeys);
            panelSource.Items.Add(btnLicense);

            var panel = new RibbonPanel { Source = panelSource };
            tab.Panels.Add(panel);
            ribbon.Tabs.Add(tab);

            tab.IsActive = true;
        }

        public static void Remove()
        {
            var ribbon = ComponentManager.Ribbon;
            if (ribbon == null) return;
            for (int i = ribbon.Tabs.Count - 1; i >= 0; i--)
            {
                if (ribbon.Tabs[i].Id == TabId) ribbon.Tabs.RemoveAt(i);
            }
        }

        // ---------- Icon drawing ----------

        private static BitmapSource DrawAutoIcon(int size)
        {
            return BuildBitmap(size, g =>
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // Outer dashed "detection" frame
                var dashPen = new Pen(Color.FromArgb(255, 40, 150, 60), Math.Max(1f, size / 14f))
                {
                    DashStyle = DashStyle.Dash
                };
                float pad = size * 0.1f;
                g.DrawRectangle(dashPen, pad, pad, size - pad * 2, size - pad * 2);

                // Magic-wand star sparkle at top-right hinting at "auto"
                var starBrush = new SolidBrush(Color.FromArgb(255, 200, 140, 30));
                float sx = size * 0.72f, sy = size * 0.28f;
                float sh = size * 0.2f;
                g.FillPolygon(starBrush, new[]
                {
                    new PointF(sx, sy - sh),
                    new PointF(sx + sh * 0.25f, sy - sh * 0.25f),
                    new PointF(sx + sh, sy),
                    new PointF(sx + sh * 0.25f, sy + sh * 0.25f),
                    new PointF(sx, sy + sh),
                    new PointF(sx - sh * 0.25f, sy + sh * 0.25f),
                    new PointF(sx - sh, sy),
                    new PointF(sx - sh * 0.25f, sy - sh * 0.25f),
                });

                // Inner small frame indicating the generated block
                var innerPen = new Pen(Color.FromArgb(255, 60, 110, 180), Math.Max(1f, size / 18f));
                float ip = size * 0.32f;
                g.DrawRectangle(innerPen, ip, ip, size - ip * 2, size - ip * 2);
            });
        }

        private static BitmapSource DrawSttIcon(int size)
        {
            return BuildBitmap(size, g =>
            {
                g.Clear(Color.Transparent);

                // 3 small framed squares with numbers 1, 2, 3
                var pen = new Pen(Color.FromArgb(255, 60, 110, 180), Math.Max(1f, size / 16f));
                var font = new Font("Segoe UI", size * 0.28f, FontStyle.Bold);
                var brush = new SolidBrush(Color.FromArgb(255, 30, 70, 140));
                var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };

                int cell = (int)(size * 0.42);
                int gap = (int)(size * 0.08);
                int startX = (size - (cell * 2 + gap)) / 2;
                int startY = (size - (cell * 2 + gap)) / 2;

                for (int i = 0; i < 3; i++)
                {
                    int x = startX + (i % 2) * (cell + gap);
                    int y = startY + (i / 2) * (cell + gap);
                    g.DrawRectangle(pen, x, y, cell, cell);
                    g.DrawString((i + 1).ToString(), font, brush,
                        new RectangleF(x, y, cell, cell), sf);
                }
            });
        }

        private static BitmapSource DrawLayoutIcon(int size)
        {
            return BuildBitmap(size, g =>
            {
                g.Clear(Color.Transparent);

                float pad = size * 0.1f;
                float w = size - pad * 2;
                float h = size - pad * 2;

                // outer paper sheet
                var paperPen = new Pen(Color.FromArgb(255, 60, 60, 60), Math.Max(1f, size / 16f));
                g.DrawRectangle(paperPen, pad, pad, w, h);

                // inner viewport area
                float vpPad = size * 0.22f;
                var vpPen = new Pen(Color.FromArgb(255, 200, 80, 30), Math.Max(1f, size / 18f));
                g.DrawRectangle(vpPen, vpPad, vpPad, size - vpPad * 2, size - vpPad * 2 - size * 0.14f);

                // title block strip at bottom
                var titleBrush = new SolidBrush(Color.FromArgb(120, 60, 110, 180));
                g.FillRectangle(titleBrush, pad, size - pad - size * 0.16f, w, size * 0.16f);
            });
        }

        private static BitmapSource DrawCleanIcon(int size)
        {
            return BuildBitmap(size, g =>
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // A faded frame
                var framePen = new Pen(Color.FromArgb(140, 120, 120, 120), Math.Max(1f, size / 16f));
                float pad = size * 0.15f;
                g.DrawRectangle(framePen, pad, pad, size - pad * 2, size - pad * 2);

                // Big red X on top
                var xPen = new Pen(Color.FromArgb(255, 200, 50, 50), Math.Max(2f, size / 7f))
                {
                    StartCap = LineCap.Round,
                    EndCap = LineCap.Round
                };
                float a = size * 0.22f;
                float b = size * 0.78f;
                g.DrawLine(xPen, a, a, b, b);
                g.DrawLine(xPen, b, a, a, b);
            });
        }

        private static BitmapSource DrawPrintIcon(int size)
        {
            return BuildBitmap(size, g =>
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;

                var dark = Color.FromArgb(255, 55, 65, 75);
                var blue = Color.FromArgb(255, 55, 110, 180);
                var green = Color.FromArgb(255, 40, 150, 80);

                float pad = size * 0.12f;
                float paperX = size * 0.25f;
                float paperY = size * 0.06f;
                float paperW = size * 0.50f;
                float paperH = size * 0.38f;
                using (var paperBrush = new SolidBrush(Color.White))
                using (var paperPen = new Pen(blue, Math.Max(1f, size / 18f)))
                {
                    g.FillRectangle(paperBrush, paperX, paperY, paperW, paperH);
                    g.DrawRectangle(paperPen, paperX, paperY, paperW, paperH);
                }

                float bodyY = size * 0.36f;
                float bodyH = size * 0.34f;
                using (var bodyBrush = new SolidBrush(Color.FromArgb(255, 225, 230, 235)))
                using (var bodyPen = new Pen(dark, Math.Max(1f, size / 16f)))
                {
                    g.FillRectangle(bodyBrush, pad, bodyY, size - pad * 2, bodyH);
                    g.DrawRectangle(bodyPen, pad, bodyY, size - pad * 2, bodyH);
                }

                float outX = size * 0.23f;
                float outY = size * 0.61f;
                float outW = size * 0.54f;
                float outH = size * 0.30f;
                using (var outBrush = new SolidBrush(Color.White))
                using (var outPen = new Pen(green, Math.Max(1f, size / 18f)))
                {
                    g.FillRectangle(outBrush, outX, outY, outW, outH);
                    g.DrawRectangle(outPen, outX, outY, outW, outH);
                }

                using (var dotBrush = new SolidBrush(green))
                    g.FillEllipse(dotBrush, size * 0.68f, size * 0.45f, size * 0.10f, size * 0.10f);
            });
        }

        private static BitmapSource DrawSheetSetIcon(int size)
        {
            return BuildBitmap(size, g =>
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var blue = Color.FromArgb(255, 55, 110, 180);
                var green = Color.FromArgb(255, 45, 145, 85);
                float w = size * 0.58f;
                float h = size * 0.68f;
                float x = size * 0.16f;
                float y = size * 0.14f;
                using (var paper = new SolidBrush(Color.White))
                using (var pen = new Pen(blue, Math.Max(1f, size / 18f)))
                {
                    for (int i = 2; i >= 0; i--)
                    {
                        float d = i * size * 0.09f;
                        g.FillRectangle(paper, x + d, y - d, w, h);
                        g.DrawRectangle(pen, x + d, y - d, w, h);
                    }
                }
                using (var checkPen = new Pen(green, Math.Max(2f, size / 9f)))
                {
                    checkPen.StartCap = LineCap.Round;
                    checkPen.EndCap = LineCap.Round;
                    g.DrawLines(checkPen, new[]
                    {
                        new PointF(size * 0.28f, size * 0.58f),
                        new PointF(size * 0.43f, size * 0.72f),
                        new PointF(size * 0.75f, size * 0.38f),
                    });
                }
            });
        }

        private static BitmapSource DrawFrameSetupIcon(int size)
        {
            return BuildBitmap(size, g =>
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;

                var blue = Color.FromArgb(255, 55, 110, 180);
                var orange = Color.FromArgb(255, 200, 100, 30);
                float pad = size * 0.10f;
                float w = size - pad * 2;
                float h = size - pad * 2;

                using (var paper = new SolidBrush(Color.White))
                using (var paperPen = new Pen(Color.FromArgb(255, 60, 60, 60), Math.Max(1f, size / 16f)))
                {
                    g.FillRectangle(paper, pad, pad, w, h);
                    g.DrawRectangle(paperPen, pad, pad, w, h);
                }

                // Title-block strip at bottom with field slots
                float stripH = size * 0.28f;
                float stripY = size - pad - stripH;
                using (var strip = new SolidBrush(Color.FromArgb(40, 55, 110, 180)))
                using (var stripPen = new Pen(blue, Math.Max(1f, size / 18f)))
                {
                    g.FillRectangle(strip, pad, stripY, w, stripH);
                    g.DrawRectangle(stripPen, pad, stripY, w, stripH);
                }

                float slotH = size * 0.08f;
                float slotY = stripY + size * 0.10f;
                float slotW = size * 0.22f;
                float gap = size * 0.04f;
                float sx = pad + size * 0.06f;
                using (var slot = new SolidBrush(Color.FromArgb(220, 255, 255, 255)))
                using (var slotPen = new Pen(orange, Math.Max(1f, size / 22f)))
                {
                    for (int i = 0; i < 3; i++)
                    {
                        float x = sx + i * (slotW + gap);
                        g.FillRectangle(slot, x, slotY, slotW, slotH);
                        g.DrawRectangle(slotPen, x, slotY, slotW, slotH);
                    }
                }

                // Small gear hint (setup) at top-right of sheet
                float cx = size * 0.72f, cy = size * 0.32f, r = size * 0.14f;
                using (var gearPen = new Pen(blue, Math.Max(1.5f, size / 14f)))
                {
                    gearPen.StartCap = LineCap.Round;
                    gearPen.EndCap = LineCap.Round;
                    g.DrawEllipse(gearPen, cx - r * 0.55f, cy - r * 0.55f, r * 1.1f, r * 1.1f);
                    for (int i = 0; i < 4; i++)
                    {
                        double a = i * Math.PI / 2.0;
                        float x1 = cx + (float)(Math.Cos(a) * r * 0.55f);
                        float y1 = cy + (float)(Math.Sin(a) * r * 0.55f);
                        float x2 = cx + (float)(Math.Cos(a) * r);
                        float y2 = cy + (float)(Math.Sin(a) * r);
                        g.DrawLine(gearPen, x1, y1, x2, y2);
                    }
                }
            });
        }

        private static BitmapSource DrawLicenseIcon(int size)
        {
            return BuildBitmap(size, g =>
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // Key body (circle) + shaft + teeth — represents a license/activation key.
                var keyColor = Color.FromArgb(255, 200, 150, 30);
                var pen = new Pen(keyColor, Math.Max(1f, size / 9f))
                {
                    StartCap = LineCap.Round,
                    EndCap = LineCap.Round
                };

                float r = size * 0.18f;
                float cx = size * 0.30f;
                float cy = size * 0.30f;
                g.DrawEllipse(pen, cx - r, cy - r, r * 2, r * 2);

                float x1 = cx + r * 0.7f;
                float y1 = cy + r * 0.7f;
                float x2 = size * 0.88f;
                float y2 = size * 0.88f;
                g.DrawLine(pen, x1, y1, x2, y2);

                float tooth = size * 0.16f;
                float midX = (x1 + x2) * 0.5f;
                float midY = (y1 + y2) * 0.5f;
                g.DrawLine(pen, midX, midY, midX + tooth, midY - tooth);
                float t2x = midX + (x2 - x1) * 0.25f;
                float t2y = midY + (y2 - y1) * 0.25f;
                g.DrawLine(pen, t2x, t2y, t2x + tooth * 0.7f, t2y - tooth * 0.7f);
            });
        }

        private static BitmapSource DrawKeysIcon(int size)
        {
            return BuildBitmap(size, g =>
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // Keyboard body
                var bodyPen = new Pen(Color.FromArgb(255, 55, 65, 75), Math.Max(1f, size / 16f));
                using (var bodyBrush = new SolidBrush(Color.FromArgb(255, 230, 234, 238)))
                {
                    float x = size * 0.10f, y = size * 0.28f, w = size * 0.80f, h = size * 0.44f;
                    g.FillRectangle(bodyBrush, x, y, w, h);
                    g.DrawRectangle(bodyPen, x, y, w, h);
                }

                // Key dots (3 columns x 2 rows) + a wide space bar
                using (var key = new SolidBrush(Color.FromArgb(255, 60, 110, 180)))
                {
                    float kw = size * 0.12f, kh = size * 0.10f;
                    float sx = size * 0.18f, sy = size * 0.36f;
                    float gx = size * 0.22f, gy = size * 0.16f;
                    for (int r = 0; r < 2; r++)
                        for (int c = 0; c < 3; c++)
                            g.FillRectangle(key, sx + c * gx, sy + r * gy, kw, kh);
                }
            });
        }

        private static BitmapSource DrawResetIcon(int size)
        {
            return BuildBitmap(size, g =>
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // circular arrow (reset symbol)
                var pen = new Pen(Color.FromArgb(255, 180, 60, 60), Math.Max(1f, size / 10f))
                {
                    StartCap = LineCap.Round,
                    EndCap = LineCap.Round
                };
                float pad = size * 0.18f;
                var rect = new RectangleF(pad, pad, size - pad * 2, size - pad * 2);
                g.DrawArc(pen, rect, -45, 270);

                // arrow head
                float cx = size / 2f, cy = size / 2f;
                double ang = Math.PI * (-45) / 180.0;
                float r = (size - pad * 2) / 2f;
                float ax = cx + (float)(Math.Cos(ang) * r);
                float ay = cy + (float)(Math.Sin(ang) * r);
                var headBrush = new SolidBrush(Color.FromArgb(255, 180, 60, 60));
                float hs = size * 0.18f;
                var pts = new[]
                {
                    new PointF(ax + hs, ay),
                    new PointF(ax - hs * 0.3f, ay - hs),
                    new PointF(ax - hs * 0.3f, ay + hs),
                };
                g.FillPolygon(headBrush, pts);
            });
        }

        private static BitmapSource BuildBitmap(int size, Action<Graphics> draw)
        {
            using (var bmp = new Bitmap(size, size))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                    draw(g);
                }
                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    ms.Position = 0;
                    var img = new BitmapImage();
                    img.BeginInit();
                    img.CacheOption = BitmapCacheOption.OnLoad;
                    img.StreamSource = ms;
                    img.EndInit();
                    img.Freeze();
                    return img;
                }
            }
        }

        private class RelayCommand : System.Windows.Input.ICommand
        {
            private readonly Action _action;
            public RelayCommand(Action action) { _action = action; }
            public bool CanExecute(object parameter) => true;
            public void Execute(object parameter) => _action();
            public event EventHandler CanExecuteChanged { add { } remove { } }
        }
    }
}
