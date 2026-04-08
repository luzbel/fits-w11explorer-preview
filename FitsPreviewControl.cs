using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using System.Windows.Forms;

namespace FitsPreviewHandler
{
    // ── Image metadata extracted from the FITS primary header ───────────
    internal struct ImageInfo
    {
        public int    Width;         // NAXIS1
        public int    Height;        // NAXIS2
        public int    Planes;        // NAXIS3 (defaults to 1)
        public int    BitPix;        // BITPIX
        public double BZero;         // BZERO  (default 0)
        public double BScale;        // BSCALE (default 1)
        public string BayerPattern;  // null | "RGGB" | "BGGR" | "GRBG" | "GBRG"
        public string Filter;        // FILTER keyword value
        public long   DataOffset;    // byte offset of first pixel in file
        public bool   HasImage;      // true when Width>0 && Height>0

        // Additional fields for Property Handler
        public string Camera;
        public string Instrument;
        public string Telescope;
        public string Object;
        public string DateObs;
        public double Exposure;
        public string Software;

        public int  BytesPerPixel => Math.Abs(BitPix) / 8;
        public override string ToString() =>
            $"{Width}x{Height} Planes={Planes} BITPIX={BitPix} " +
            $"BZero={BZero} BScale={BScale} DataOffset={DataOffset}" +
            (BayerPattern != null ? $" Bayer={BayerPattern}" : "") +
            (Filter       != null ? $" Filter={Filter}"       : "");
    }

    // ── Preview UserControl ─────────────────────────────────────────────
    public class FitsPreviewControl : UserControl
    {
        // ── Shared log (same file as FitsPreviewHandlerExtension) ───────
        public static string LogPath;

        static FitsPreviewControl()
        {
            try
            {
                // LocalLow is the designated spot for low-integrity processes (prevhost)
                string userProfile = Environment.GetEnvironmentVariable("USERPROFILE");
                string logDir = Path.Combine(userProfile, "AppData", "LocalLow", "FitsPreviewHandler");
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                LogPath = Path.Combine(logDir, "fits_trace.log");
            }
            catch
            {
                // Last ditch effort: root of local temp
                LogPath = Path.Combine(Path.GetTempPath(), "fits_trace.log");
            }
        }

        internal static void Log(string msg)
        {
            if (!Settings.EnableTracing) return;
            msg = $"[{DateTime.Now:HH:mm:ss.fff}] [UI] [T{System.Threading.Thread.CurrentThread.ManagedThreadId}] {msg}";
            System.Diagnostics.Debug.WriteLine(msg);
            try
            {
                File.AppendAllText(LogPath, msg + "\n");
            }
            catch { /* fallback only */ }
        }

        // ── Controls ────────────────────────────────────────────────────
        private Panel         _topPanel;
        private Label         _lblTitle;
        private SplitContainer _split;
        private DataGridView  _gridHeader;
        private PictureBox    _pictureBox;
        private Label         _lblProgress;
        private Label         _lblImageHint;
        private Label         _lblLogStatus;
        private TextBox       _txtError;

        // ── Concurrency ─────────────────────────────────────────────────
        private System.Threading.CancellationTokenSource _cts;
        private readonly object _syncRoot = new object();

        private void CancelOldLoad()
        {
            lock (_syncRoot)
            {
                if (_cts != null)
                {
                    _cts.Cancel();
                    _cts.Dispose();
                    _cts = null;
                }
                _cts = new System.Threading.CancellationTokenSource();
            }
        }

        public FitsPreviewControl()
        {
            Log("FitsPreviewControl.ctor");
            try   { InitializeComponent(); Log("ctor — OK"); }
            catch (Exception ex) { Log("ctor — EXCEPTION: " + ex); throw; }
        }

        private void InitializeComponent()
        {
            BackColor = Color.FromArgb(24, 24, 36);
            ForeColor = Color.FromArgb(220, 220, 235);
            AutoScroll = false; // Disable global scrollbar to use inner panel scrolling instead

            // ── Title bar ────────────────────────────────────────────────
            _topPanel = new Panel
            {
                Dock = DockStyle.Top, Height = 36,
                BackColor = Color.FromArgb(40, 40, 60),
                Padding = new Padding(8, 0, 0, 0)
            };
            _lblTitle = new Label
            {
                AutoSize = false, Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "FITS Preview",
                ForeColor = Color.FromArgb(140, 200, 255),
                Font = new Font("Segoe UI", 10f, FontStyle.Bold)
            };
            _lblLogStatus = new Label
            {
                AutoSize = true, Dock = DockStyle.Right,
                TextAlign = ContentAlignment.MiddleRight,
                Text = "Log: ?",
                ForeColor = Color.FromArgb(120, 120, 150),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic),
                Padding = new Padding(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            _topPanel.Controls.Add(_lblTitle);
            _topPanel.Controls.Add(_lblLogStatus);

            // ── SplitContainer (Top=Image, Bottom=Grid) ──────────────────
            _split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                FixedPanel = FixedPanel.Panel1,
                BorderStyle = BorderStyle.None,
                Panel1MinSize = 250, Panel2MinSize = 50,
                BackColor = Color.FromArgb(24, 24, 36)
            };
            
            // Set initial position: handle initialization and the 50/50 split during resizing
            _split.Resize += (s, e) => {
                if (_split.Height < 100) return; // Wait for valid size

                int savedPos = Settings.SplitterDistance;
                if (savedPos > 0)
                {
                    // Enforce our 250px minimum safety even for saved values
                    int targetPos = Math.Max(250, savedPos);
                    if (_split.SplitterDistance != targetPos && targetPos < _split.Height - 50)
                    {
                        _split.SplitterDistance = targetPos;
                        Log($"Resize (Stored) — SplitterDistance forced to {targetPos} (Height={_split.Height})");
                    }
                }
                else if (savedPos == -1)
                {
                    // Dynamic 50/50 split
                    _split.SplitterDistance = _split.Height / 2;
                    Log($"Resize (Initial Auto) — SplitterDistance set to {_split.SplitterDistance} (Height={_split.Height})");
                }
            };

            // Persist splitter position
            _split.SplitterMoved += (s, e) => Settings.SplitterDistance = _split.SplitterDistance;
            
            // Image Panel (Panel1)
            _pictureBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.FromArgb(18, 18, 28)
            };

            _lblProgress = new Label
            {
                AutoSize = true,
                Text = "",
                ForeColor = Color.FromArgb(180, 190, 220),
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };
            
            // Handle centering of progress label
            _split.Panel1.Resize += (s, e) => {
                _lblProgress.Left = (_split.Panel1.Width - _lblProgress.Width) / 2;
                _lblProgress.Top = (_split.Panel1.Height - _lblProgress.Height) / 2;
            };
            _lblProgress.SizeChanged += (s, e) => {
                _lblProgress.Left = (_split.Panel1.Width - _lblProgress.Width) / 2;
                _lblProgress.Top = (_split.Panel1.Height - _lblProgress.Height) / 2;
            };

            _lblImageHint = new Label
            {
                AutoSize = false, Dock = DockStyle.Bottom, Height = 40,
                TextAlign = ContentAlignment.MiddleCenter,
                Text = "",
                ForeColor = Color.FromArgb(110, 115, 140),
                BackColor = Color.FromArgb(15, 15, 25),
                Font = new Font("Consolas", 8.5f),
                Visible = false
            };

            _split.Panel1.Controls.Add(_lblProgress);
            _split.Panel1.Controls.Add(_lblImageHint);
            _split.Panel1.Controls.Add(_pictureBox);

            // Grid (Panel2)
            _gridHeader = new DataGridView
            {
                Dock = DockStyle.Fill,
                RowHeadersVisible = false, 
                AllowUserToAddRows = false,
                AllowUserToResizeRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                BackgroundColor = Color.FromArgb(28, 28, 42),
                GridColor = Color.FromArgb(50, 50, 70),
                BorderStyle = BorderStyle.None,
                ScrollBars = ScrollBars.Both, 
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(28, 28, 42),
                    ForeColor = Color.FromArgb(210, 215, 235),
                    SelectionBackColor = Color.FromArgb(60, 80, 130),
                    SelectionForeColor = Color.White,
                    Font = new Font("Consolas", 9f),
                    Padding = new Padding(4, 2, 4, 2)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(40, 40, 60),
                    ForeColor = Color.FromArgb(140, 200, 255),
                    Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                    Padding = new Padding(4, 4, 4, 4)
                },
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 28,
                RowTemplate = { Height = 22 },
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            };
            _gridHeader.Columns.Add(new DataGridViewTextBoxColumn
                { Name="Keyword", HeaderText="Keyword", Width=100, MinimumWidth=50, Resizable=DataGridViewTriState.True });
            _gridHeader.Columns.Add(new DataGridViewTextBoxColumn
                { Name="Value", HeaderText="Value", Width=150, MinimumWidth=60, Resizable=DataGridViewTriState.True });
            _gridHeader.Columns.Add(new DataGridViewTextBoxColumn
                { Name="Comment", HeaderText="Comment",
                  AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill, MinimumWidth=60, Resizable=DataGridViewTriState.True });
            
            _gridHeader.RowsAdded += (s, e) => {
                for (int i = e.RowIndex; i < e.RowIndex + e.RowCount && i < _gridHeader.Rows.Count; i++)
                    _gridHeader.Rows[i].DefaultCellStyle.BackColor =
                        (i % 2 == 0) ? Color.FromArgb(28, 28, 42) : Color.FromArgb(33, 33, 50);
            };

            _split.Panel2.Controls.Add(_gridHeader);

            // ── Error box (hidden) ──────────────────────────────────────
            _txtError = new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, Visible = false,
                BackColor = Color.FromArgb(60, 20, 20),
                ForeColor = Color.FromArgb(255, 160, 160),
                Font = new Font("Consolas", 9f), ScrollBars = ScrollBars.Vertical
            };

            // ── Context Menu ──────────────────────────────────────────────
            var ctxMenu = new ContextMenuStrip();
            ctxMenu.RenderMode = ToolStripRenderMode.System;
            
            this.ContextMenuStrip = ctxMenu;
            _pictureBox.ContextMenuStrip = ctxMenu;
            _gridHeader.ContextMenuStrip = ctxMenu;
            _split.ContextMenuStrip = ctxMenu;
            _split.Panel1.ContextMenuStrip = ctxMenu;
            _split.Panel2.ContextMenuStrip = ctxMenu;
            _topPanel.ContextMenuStrip = ctxMenu;
            _lblImageHint.ContextMenuStrip = ctxMenu;

            ctxMenu.Opening += (s, e) => {
                ctxMenu.Items.Clear();
                bool currentShowImg = Settings.ShowImage;
                bool currentLogOn = Settings.EnableTracing;

                var itemImg = new ToolStripMenuItem(currentShowImg ? "Ocultar imagen (carga más rápida). Solo mostrar metadatos" : "Previsualizar imagen con auto-stretch");
                itemImg.Click += (sender, args) => {
                    Settings.ShowImage = !currentShowImg;
                    _lblImageHint.Text = "¡Guardado! Selecciona otro archivo FITS para previsualizar con la nueva configuración.";
                    _lblImageHint.ForeColor = Color.FromArgb(180, 255, 180);
                    _lblImageHint.Visible = true;
                };
                
                var itemLog = new ToolStripMenuItem(currentLogOn ? "Desactivar trazas (Trace: OFF)" : "Activar trazas para depurar en %USERPROFILE%\\AppData\\LocalLow\\FitsPreviewHandler  (Trace: ON)");
                itemLog.Click += (sender, args) => {
                    Settings.EnableTracing = !currentLogOn;
                    _lblImageHint.Text = "¡Guardado! Selecciona otro archivo FITS para aplicar la nueva configuración de trazas.";
                    _lblImageHint.ForeColor = Color.FromArgb(180, 255, 180);
                    _lblImageHint.Visible = true;
                };

                ctxMenu.Items.Add(itemImg);
                ctxMenu.Items.Add(itemLog);
                
                ctxMenu.Items.Add(new ToolStripSeparator());

                if (_pictureBox.Image != null && currentShowImg)
                {
                    var itemCopyImg = new ToolStripMenuItem($"Copiar imagen ({_pictureBox.Image.Width}x{_pictureBox.Image.Height})");
                    itemCopyImg.Click += (sender, args) => {
                        try { Clipboard.SetImage(_pictureBox.Image); } catch { }
                    };
                    ctxMenu.Items.Add(itemCopyImg);
                }
                
                if (_gridHeader.SelectedRows.Count > 0)
                {
                    var itemCopy = new ToolStripMenuItem("Copiar fila seleccionada");
                    itemCopy.Click += (sender, args) => {
                        var row = _gridHeader.SelectedRows[0];
                        string text = $"{row.Cells[0].Value}={row.Cells[1].Value} // {row.Cells[2].Value}";
                        try { Clipboard.SetText(text); } catch { }
                    };
                    ctxMenu.Items.Add(itemCopy);
                }

                var itemCopyCsv = new ToolStripMenuItem("Copiar toda la tabla (CSV)");
                itemCopyCsv.Click += (sender, args) => {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("\"Keyword\",\"Value\",\"Comment\"");
                    foreach (DataGridViewRow r in _gridHeader.Rows)
                    {
                        string c0 = r.Cells[0].Value?.ToString() ?? "";
                        string c1 = r.Cells[1].Value?.ToString() ?? "";
                        string c2 = r.Cells[2].Value?.ToString() ?? "";
                        c0 = c0.Replace("\"", "\"\"");
                        c1 = c1.Replace("\"", "\"\"");
                        c2 = c2.Replace("\"", "\"\"");
                        sb.AppendLine($"\"{c0}\",\"{c1}\",\"{c2}\"");
                    }
                    try { Clipboard.SetText(sb.ToString()); } catch { }
                };
                ctxMenu.Items.Add(itemCopyCsv);
            };

            Controls.Add(_split);
            Controls.Add(_topPanel);
            Controls.Add(_txtError);
            Log("InitializeComponent — done");
        }

        // ── Public entry points ─────────────────────────────────────────
        public void LoadFits(string filePath)
        {
            Log($"LoadFits(string) — '{filePath}'");
            try
            {
                if (!File.Exists(filePath))
                {
                    Log("LoadFits — file NOT FOUND");
                    ShowError($"File not found:\n{filePath}");
                    return;
                }
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    LoadFits(fs, Path.GetFileName(filePath));
                }
            }
            catch (Exception ex)
            {
                Log("LoadFits(string) — EXCEPTION: " + ex);
                ShowError("Error opening FITS file:\r\n" + ex.Message);
            }
        }

        public void LoadFits(Stream stream, string fileName)
        {
            Log($"LoadFits(Stream) — '{fileName}' length={stream.Length}");
            
            bool showImg = Settings.ShowImage;
            bool logOn   = Settings.EnableTracing;
            Log($"LoadFits — ShowImage={showImg} Log={logOn} SplitterDistance={_split.SplitterDistance} TotalHeight={_split.Height}");

            Action updateLayout = () => {
                // If image is disabled, show placeholder text in Panel1 instead of collapsing it
                _pictureBox.Visible = showImg;
                _lblImageHint.Visible = true;
                
                if (!showImg)
                {
                    _lblProgress.Visible = false;
                }
                
                _lblImageHint.Text = "Right-Click for configuration and copy options";
                _lblImageHint.Visible = true;

                _lblLogStatus.Text = "Trace: " + (logOn ? "ON" : "OFF");
                _lblLogStatus.ForeColor = logOn ? Color.FromArgb(150, 255, 150) : Color.FromArgb(120, 120, 140);
            };
            if (InvokeRequired) Invoke(updateLayout); else updateLayout();

            try
            {
                var (rows, info) = ParseFitsStream(stream);
                Log($"LoadFits — {rows.Count} keywords  image={info}");

                Action populate = () => PopulateGrid(rows, fileName, info);
                if (InvokeRequired) Invoke(populate); else populate();

                if (showImg && info.HasImage)
                {
                    CancelOldLoad();
                    StartImageLoad(stream, info, _cts.Token);
                }
                else if (!showImg)
                    Log("LoadFits — image loading disabled in registry");
                else
                    Log("LoadFits — no image data to display");
            }
            catch (Exception ex)
            {
                Log("LoadFits(Stream) — EXCEPTION: " + ex);
                ShowError("Error reading FITS stream:\r\n" + ex.Message);
            }
        }

        // ── Grid population ─────────────────────────────────────────────
        private void PopulateGrid(
            List<(string kw, string val, string comment)> rows,
            string fileName, ImageInfo info)
        {
            Log($"PopulateGrid — {rows.Count} rows");
            _gridHeader.Rows.Clear();

            string imgNote = info.HasImage
                ? $" · {info.Width}×{info.Height}" +
                  (info.Planes > 1 ? $"×{info.Planes}" : "") +
                  $"  BITPIX={info.BitPix}"
                : " · no image";
            
            if (!Settings.ShowImage) imgNote += " (Image hidden)";
            _lblTitle.Text = fileName + imgNote;

            foreach (var (kw, val, comment) in rows)
            {
                int idx = _gridHeader.Rows.Add(kw, val, comment);
                if      (kw == "END")                       _gridHeader.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(255, 120, 120);
                else if (kw == "COMMENT" || kw == "HISTORY") _gridHeader.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(160, 200, 140);
                else if (kw.StartsWith("NAXIS"))             _gridHeader.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(255, 200, 100);
            }
            Log("PopulateGrid — done");
        }

        // ── Async image load ────────────────────────────────────────────
        private void StartImageLoad(Stream stream, ImageInfo info, System.Threading.CancellationToken token)
        {
            Log($"StartImageLoad — {info.Width}×{info.Height} planes={info.Planes} bitpix={info.BitPix}");

            if (info.BitPix == 0)
            {
                BeginInvoke(new Action(() => _lblProgress.Text = "⚠ BITPIX=0, cannot render"));
                return;
            }

            const long MAX_BYTES = 2L * 1024 * 1024 * 1024;
            long dataBytes = (long)info.Width * info.Height * info.Planes * info.BytesPerPixel;
            if (dataBytes > MAX_BYTES)
            {
                BeginInvoke(new Action(() => _lblProgress.Text = $"⚠ Image too large ({dataBytes/1024/1024} MB)"));
                return;
            }

            Task.Run(() =>
            {
                try
                {
                    void Report(string msg) => BeginInvoke(new Action(() => {
                        _lblProgress.Text = msg;
                        _lblProgress.Visible = true;
                    }));

                    Bitmap bmp = RenderImage(stream, info, token, Report);
                    if (bmp != null && !token.IsCancellationRequested)
                    {
                        BeginInvoke(new Action(() =>
                        {
                            _pictureBox.Image = bmp;
                            _pictureBox.Visible = true;
                            _lblProgress.Visible = false;
                        }));
                    }
                }
                catch (Exception ex)
                {
                    Log("StartImageLoad — EXCEPTION: " + ex);
                    BeginInvoke(new Action(() => {
                        _lblProgress.Text = "Render Error: " + ex.Message;
                        _lblProgress.Visible = true;
                    }));
                }
            }, token);
        }

        // ── Rendering pipeline ──────────────────────────────────────────
        private static Bitmap RenderImage(Stream stream, ImageInfo info, System.Threading.CancellationToken token, Action<string> reportProgress = null)
        {
            const int MAX_DIM = 2000; // Increased max dim for better resolution in large layouts
            double aspect = (double)info.Width / info.Height;
            int outW, outH;
            if (info.Width >= info.Height)
            { outW = Math.Min(info.Width, MAX_DIM); outH = Math.Max(1, (int)(outW / aspect)); }
            else
            { outH = Math.Min(info.Height, MAX_DIM); outW = Math.Max(1, (int)(outH * aspect)); }

            Log($"RenderImage — target {outW}×{outH}");

            // Read color/mono data
            float[][] planes = ReadAllPlanes(stream, info, outW, outH, out int aW, out int aH, token, reportProgress);
            if (token.IsCancellationRequested) return null;
            Log($"RenderImage — planes={planes.Length} actual={aW}×{aH}");

            if (planes.Length >= 3)
                return PlanesToColorBitmap(planes[0], planes[1], planes[2], aW, aH, token);

            // Mono/Narrowband
            AdaptiveStretch(planes[0], out float low, out float high);
            if (token.IsCancellationRequested) return null;
            Log($"RenderImage — stretch low={low:G4} high={high:G4}");

            Color tint = GetFilterTint(info.Filter);
            return PixelsToBitmap(planes[0], aW, aH, low, high, tint);
        }

        private static float[][] ReadAllPlanes(
            Stream stream, ImageInfo info,
            int targetW, int targetH,
            out int actualW, out int actualH,
            System.Threading.CancellationToken token,
            Action<string> reportProgress = null)
        {
            // If Bayer, we will produce 3 planes (RGB) directly during downsampling
            bool isBayer = !string.IsNullOrEmpty(info.BayerPattern) && info.Planes == 1;

            int sx = Math.Max(1, info.Width  / targetW);
            int sy = Math.Max(1, info.Height / targetH);

            // For Bayer, we MUST ensure step is even (2x2 cells)
            if (isBayer) {
                sx = (sx < 2) ? 2 : (sx % 2 != 0 ? sx + 1 : sx);
                sy = (sy < 2) ? 2 : (sy % 2 != 0 ? sy + 1 : sy);
            }

            actualW = 0; for (int x = 0; x <= info.Width - (isBayer ? 2 : 1);  x += sx) actualW++;
            actualH = 0; for (int y = 0; y <= info.Height - (isBayer ? 2 : 1); y += sy) actualH++;
            Log($"ReadAllPlanes — {info.Width}x{info.Height} stride {sx}x{sy} -> {actualW}x{actualH} (Bayer:{isBayer})");

            int bpp = info.BytesPerPixel;
            long rowBytes = (long)info.Width * bpp;
            
            int outPlanes = isBayer ? 3 : info.Planes;
            float[][] planes = new float[outPlanes][];
            for (int p = 0; p < outPlanes; p++) planes[p] = new float[actualW * actualH];

            byte[] rowBuf = new byte[rowBytes];
            byte[] nextRowBuf = isBayer ? new byte[rowBytes] : null;

            for (int p = 0; p < info.Planes && p < (isBayer ? 1 : outPlanes); p++)
            {
                long planeOff = info.DataOffset + (long)p * info.Height * rowBytes;
                int outY = 0;
                for (int y = 0; y <= info.Height - (isBayer ? 2 : 1) && outY < actualH; y += sy)
                {
                    if (token.IsCancellationRequested) break;

                    if (y % (sy * 20) == 0) // Report progress every 20 output rows
                    {
                        double progress = 100.0 * (p * info.Height + y) / (info.Planes * info.Height);
                        reportProgress?.Invoke($"Reading FITS data... {progress:F0}%");
                    }

                    stream.Seek(planeOff + (long)y * rowBytes, SeekOrigin.Begin);
                    stream.Read(rowBuf, 0, (int)rowBytes);
                    
                    if (isBayer) {
                        stream.Seek(planeOff + (long)(y+1) * rowBytes, SeekOrigin.Begin);
                        stream.Read(nextRowBuf, 0, (int)rowBytes);
                    }

                    int outX = 0;
                    for (int x = 0; x <= info.Width - (isBayer ? 2 : 1) && outX < actualW; x += sx)
                    {
                        int idx = (actualH - 1 - outY) * actualW + outX;

                        if (isBayer) {
                            double p00 = ReadRaw(rowBuf, x * bpp, info.BitPix);
                            double p10 = ReadRaw(rowBuf, (x+1) * bpp, info.BitPix);
                            double p01 = ReadRaw(nextRowBuf, x * bpp, info.BitPix);
                            double p11 = ReadRaw(nextRowBuf, (x+1) * bpp, info.BitPix);
                            
                            float r, g, b;
                            string pat = info.BayerPattern.ToUpper();
                            if (pat == "RGGB") {
                                r=(float)p00; g=(float)(p10+p01)/2f; b=(float)p11;
                            } else if (pat == "GRBG") {
                                g=(float)(p00+p11)/2f; r=(float)p10; b=(float)p01;
                            } else if (pat == "GBRG") {
                                g=(float)(p00+p11)/2f; b=(float)p10; r=(float)p01;
                            } else { // BGGR
                                b=(float)p00; g=(float)(p10+p01)/2f; r=(float)p11; 
                            }

                            planes[0][idx] = (float)(r * info.BScale + info.BZero);
                            planes[1][idx] = (float)(g * info.BScale + info.BZero);
                            planes[2][idx] = (float)(b * info.BScale + info.BZero);
                        }
                        else {
                            double raw = ReadRaw(rowBuf, x * bpp, info.BitPix);
                            planes[p][idx] = (float)(raw * info.BScale + info.BZero);
                        }
                        outX++;
                    }
                    outY++;
                }
            }
            return planes;
        }

        // ── Big-endian pixel reader (all BITPIX variants) ───────────────
        private static double ReadRaw(byte[] buf, int off, int bitpix)
        {
            switch (bitpix)
            {
                case  8: return buf[off];
                case 16: return (short)((buf[off] << 8) | buf[off + 1]);
                case 32: return (int)(((uint)buf[off] << 24) | ((uint)buf[off+1] << 16)
                                     | ((uint)buf[off+2] << 8) | buf[off+3]);
                case -32:
                {
                    byte[] b = { buf[off+3], buf[off+2], buf[off+1], buf[off] };
                    return BitConverter.ToSingle(b, 0);
                }
                case -64:
                {
                    byte[] b = { buf[off+7], buf[off+6], buf[off+5], buf[off+4],
                                 buf[off+3], buf[off+2], buf[off+1], buf[off] };
                    return BitConverter.ToDouble(b, 0);
                }
                default: return 0;
            }
        }

        // ── Adaptive Stretch (Robust Median/MAD based) ─────────────────
        private static void AdaptiveStretch(float[] pixels, out float low, out float high)
        {
            if (pixels.Length == 0) { low = 0; high = 1; return; }
            
            // Sample for performance
            int n = pixels.Length;
            int step = Math.Max(1, n / 10000); // 10k samples max
            var sample = new List<float>();
            for(int i=0; i<n; i+=step) sample.Add(pixels[i]);
            var s = sample.ToArray();
            Array.Sort(s);

            float median = s[s.Length / 2];
            
            // Standard Percentile as fallback
            int loIdx = (int)(s.Length * 0.005);
            int hiIdx = (int)(s.Length * 0.995);
            low = s[loIdx];
            high = s[hiIdx];

            // If it's a deep astro image, the median is VERY low
            // Stretch to show details above background
            float background = median;
            float spread = high - background;
            
            // For Phase 1 we keep it simple but better than raw min/max
            Log($"AdaptiveStretch — med={median:G4} hi={high:G4} range={high-low:G4}");
        }

        private static Color GetFilterTint(string filter)
        {
            if (string.IsNullOrEmpty(filter)) return Color.White;
            string f = filter.ToLower();
            if (f.Contains("ha") || f.Contains("alpha")) return Color.FromArgb(255, 120, 120);
            if (f.Contains("oiii")) return Color.FromArgb(120, 255, 255);
            if (f.Contains("sii")) return Color.FromArgb(255, 100, 100);
            if (f.Contains("hb")) return Color.FromArgb(150, 150, 255);
            return Color.White;
        }

        // ── Pixels to Bitmap (with Optional Tint) ─────────────────────
        private static Bitmap PixelsToBitmap(float[] px, int w, int h, float low, float high, Color tint)
        {
            float range = high - low; if (range == 0) range = 1;
            var bmp = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            BitmapData bd = bmp.LockBits(new Rectangle(0, 0, w, h),
                                          ImageLockMode.WriteOnly,
                                          PixelFormat.Format24bppRgb);
            byte[] buf = new byte[Math.Abs(bd.Stride) * h];
            int stride = bd.Stride;
            
            float tr = tint.R / 255f;
            float tg = tint.G / 255f;
            float tb = tint.B / 255f;

            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                {
                    float norm = Math.Max(0f, Math.Min(1f, (px[y * w + x] - low) / range));
                    int i = y * stride + x * 3;
                    buf[i]   = (byte)(norm * tb * 255f);
                    buf[i+1] = (byte)(norm * tg * 255f);
                    buf[i+2] = (byte)(norm * tr * 255f);
                }
            Marshal.Copy(buf, 0, bd.Scan0, buf.Length);
            bmp.UnlockBits(bd);
            return bmp;
        }

        // ── Color bitmap from 3 planes (NAXIS3=3) ──────────────────────
        private static Bitmap PlanesToColorBitmap(
            float[] r, float[] g, float[] b, int w, int h, System.Threading.CancellationToken token)
        {
            AdaptiveStretch(r, out float lr, out float hr);
            if (token.IsCancellationRequested) return null;
            AdaptiveStretch(g, out float lg, out float hg);
            if (token.IsCancellationRequested) return null;
            AdaptiveStretch(b, out float lb, out float hb);
            if (token.IsCancellationRequested) return null;
            float rr = hr - lr; if (rr == 0) rr = 1;
            float rg = hg - lg; if (rg == 0) rg = 1;
            float rb = hb - lb; if (rb == 0) rb = 1;

            var bmp = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            BitmapData bd = bmp.LockBits(new Rectangle(0, 0, w, h),
                                          ImageLockMode.WriteOnly,
                                          PixelFormat.Format24bppRgb);
            byte[] buf = new byte[Math.Abs(bd.Stride) * h];
            int stride = bd.Stride;
            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                {
                    int pi = y * w + x;
                    byte rv = (byte)(Math.Max(0f, Math.Min(1f, (r[pi]-lr)/rr)) * 255f);
                    byte gv = (byte)(Math.Max(0f, Math.Min(1f, (g[pi]-lg)/rg)) * 255f);
                    byte bv = (byte)(Math.Max(0f, Math.Min(1f, (b[pi]-lb)/rb)) * 255f);
                    int i = y * stride + x * 3;
                    buf[i] = bv; buf[i+1] = gv; buf[i+2] = rv; // GDI+ = BGR
                }
            Marshal.Copy(buf, 0, bd.Scan0, buf.Length);
            bmp.UnlockBits(bd);
            return bmp;
        }

        // ── Bilinear Debayer ───────────────────────────────────────────
        private static Bitmap BilinearDebayer(float[] px, int w, int h, float low, float high, string pattern, System.Threading.CancellationToken token)
        {
            float range = high - low; if (range == 0) range = 1;
            var bmp = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            BitmapData bd = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
            byte[] buf = new byte[Math.Abs(bd.Stride) * h];
            int stride = bd.Stride;

            // Pattern offsets (R, G, B channels)
            // Default assumes RGGB:
            // R G
            // G B
            int rX = 0, rY = 0, bX = 1, bY = 1;
            string p = pattern.ToUpper();
            if (p == "BGGR") { rX=1; rY=1; bX=0; bY=0; }
            else if (p == "GBRG") { rX=0; rY=1; bX=1; bY=0; }
            else if (p == "GRBG") { rX=1; rY=0; bX=0; bY=1; }

            for (int y = 0; y < h; y++)
            {
                if (token.IsCancellationRequested) break;
                for (int x = 0; x < w; x++)
                {
                    float r = 0, g = 0, b = 0;
                    bool isR = (x % 2 == rX) && (y % 2 == rY);
                    bool isB = (x % 2 == bX) && (y % 2 == bY);
                    bool isG = !isR && !isB;

                    if (isR)
                    {
                        r = px[y * w + x];
                        g = GetAvg(px, w, h, x, y, 1, 0, 0, 1); // neighbors
                        b = GetAvg(px, w, h, x, y, 1, 1);       // diagonals
                    }
                    else if (isB)
                    {
                        b = px[y * w + x];
                        g = GetAvg(px, w, h, x, y, 1, 0, 0, 1);
                        r = GetAvg(px, w, h, x, y, 1, 1);
                    }
                    else // Green
                    {
                        g = px[y * w + x];
                        // If we are on a row with R, neighbors in row are R, neighbors in col are B
                        if (y % 2 == rY) { 
                            r = GetAvg(px, w, h, x, y, 1, 0); b = GetAvg(px, w, h, x, y, 0, 1);
                        } else { 
                            b = GetAvg(px, w, h, x, y, 1, 0); r = GetAvg(px, w, h, x, y, 0, 1);
                        }
                    }

                    int i = y * stride + x * 3;
                    buf[i]   = (byte)(Math.Max(0f, Math.Min(1f, (b - low) / range)) * 255f);
                    buf[i+1] = (byte)(Math.Max(0f, Math.Min(1f, (g - low) / range)) * 255f);
                    buf[i+2] = (byte)(Math.Max(0f, Math.Min(1f, (r - low) / range)) * 255f);
                }
            }

            Marshal.Copy(buf, 0, bd.Scan0, buf.Length);
            bmp.UnlockBits(bd);
            return bmp;
        }

        private static float GetAvg(float[] px, int w, int h, int x, int y, params int[] offsets)
        {
            float sum = 0; int count = 0;
            for (int i = 0; i < offsets.Length; i += 2)
            {
                int dx = offsets[i], dy = offsets[i + 1];
                // Check symmetric neighbors (+dx,+dy) and (-dx,-dy). 
                // Note: for bayer filters, step is always 1 in dx or dy.
                int[] signs = { -1, 1 };
                foreach(int s in signs) {
                    int nx = x + dx * s, ny = y + dy * s;
                    if (nx >= 0 && nx < w && ny >= 0 && ny < h) {
                        sum += px[ny * w + nx]; count++;
                    }
                }
            }
            return count > 0 ? sum / count : px[y * w + x];
        }

        // ── FITS header + image-metadata parser ─────────────────────────
        internal static (List<(string, string, string)>, ImageInfo) ParseFitsStream(Stream stream)
        {
            const int REC = 80, BLOCK = 2880;
            var result = new List<(string, string, string)>();
            var info   = new ImageInfo { BScale = 1.0, Planes = 1 };

            Log($"ParseFitsStream — pos={stream.Position}");
            stream.Seek(0, SeekOrigin.Begin);
            using (var br = new BinaryReader(stream, Encoding.ASCII, true)) // leaveOpen=true
            {
                bool endFound = false;
                int  blockIdx = 0;
                while (!endFound)
                {
                    byte[] block = br.ReadBytes(BLOCK);
                    if (block.Length < BLOCK) break;
                    int recs = block.Length / REC;

                    for (int r = 0; r < recs && !endFound; r++)
                    {
                        string rec = Encoding.ASCII.GetString(block, r * REC, REC);
                        string kw  = rec.Substring(0, 8).TrimEnd();

                        if (kw == "END")
                        {
                            result.Add(("END", "", ""));
                            endFound = true;
                            // data starts at the next 2880-byte boundary
                            info.DataOffset = (long)(blockIdx + 1) * BLOCK;
                            Log($"ParseFitsStream — END at block {blockIdx} r={r}, DataOffset={info.DataOffset}");
                            break;
                        }

                        if (string.IsNullOrWhiteSpace(kw)) continue;

                        string rest = rec.Substring(8);
                        string val = "", comment = "";
                        if (rest.Length >= 2 && rest[0] == '=' && rest[1] == ' ')
                        {
                            ParseValueComment(rest.Substring(2).TrimStart(), out val, out comment);
                            result.Add((kw, val, comment.Trim()));

                            // collect image metadata
                            switch (kw)
                            {
                                case "NAXIS1":  int.TryParse(val, out info.Width);  break;
                                case "NAXIS2":  int.TryParse(val, out info.Height); break;
                                case "NAXIS3":  int.TryParse(val, out info.Planes); break;
                                case "BITPIX":  int.TryParse(val, out info.BitPix); break;
                                case "BZERO":   double.TryParse(val,
                                    System.Globalization.NumberStyles.Float,
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    out info.BZero);  break;
                                case "BSCALE":  double.TryParse(val,
                                    System.Globalization.NumberStyles.Float,
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    out info.BScale); break;
                                case "BAYERPAT": info.BayerPattern = val.Trim(); break;
                                case "COLORTYP":
                                    if (info.BayerPattern == null) info.BayerPattern = val.Trim();
                                    break;
                                case "FILTER":  info.Filter = val.Trim(); break;
                                case "INSTRUME": info.Instrument = val; break;
                                case "TELESCOP": info.Telescope = val; break;
                                case "OBJECT": info.Object = val; break;
                                case "DATE-OBS": info.DateObs = val; break;
                                case "EXPTIME":
                                case "EXPOSURE": double.TryParse(val, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out info.Exposure); break;
                                case "SWCREATE":
                                case "CREATOR":
                                case "PROGRAM": info.Software = val; break;
                                case "CAMERA": info.Camera = val; break;
                            }
                        }
                        else
                        {
                            result.Add((kw, "", rest.Trim()));
                        }
                    }
                    blockIdx++;
                }
            }

            // Defaults
            if (info.Planes < 1)  info.Planes  = 1;
            if (info.BScale == 0) info.BScale   = 1.0;
            info.HasImage = info.Width > 0 && info.Height > 0 && info.BitPix != 0;

            Log($"ParseFitsFile — {result.Count} keywords  info={info}");
            return (result, info);
        }

        // ── FITS value/comment splitter ─────────────────────────────────
        private static void ParseValueComment(string area, out string value, out string comment)
        {
            value = ""; comment = "";
            if (area.Length == 0) return;

            if (area[0] == '\'')
            {
                var sb = new StringBuilder();
                int i = 1;
                while (i < area.Length)
                {
                    if (area[i] == '\'')
                    {
                        if (i + 1 < area.Length && area[i + 1] == '\'') { sb.Append('\''); i += 2; }
                        else { i++; break; }
                    }
                    else { sb.Append(area[i]); i++; }
                }
                value = sb.ToString().TrimEnd();
                string rest = i < area.Length ? area.Substring(i) : "";
                int sl = rest.IndexOf('/');
                comment = sl >= 0 ? rest.Substring(sl + 1).Trim() : "";
            }
            else
            {
                int sl = area.IndexOf('/');
                if (sl >= 0) { value = area.Substring(0, sl).Trim(); comment = area.Substring(sl + 1).Trim(); }
                else           value = area.Trim();
            }
        }

        public void ShowError(string message)
        {
            Log("ShowError — " + message.Replace("\r\n", " | ").Replace("\n", " | "));
            void Show()
            {
                _split.Visible      = false;
                _txtError.Visible   = true;
                _txtError.Text      = message;
            }
            if (InvokeRequired) BeginInvoke(new Action(Show)); else Show();
        }
    }
}
