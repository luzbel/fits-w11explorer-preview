using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using System.Windows.Forms;

namespace FitsPreviewHandler
{
    // ── SampledStream ───────────────────────────────────────────────
    // A minimal memory stream that only stores specific blocks (header + rows)
    // serving zeros for everything else. Absolute seeks work as in full file.
    internal class SampledStream : Stream
    {
        private readonly Dictionary<long, byte[]> _samples = new Dictionary<long, byte[]>();
        private readonly long _length;
        private long _pos;
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _length;
        public override long Position { get { return _pos; } set { _pos = value; } }
        public long BufferedBytes { get; private set; }
        public int SampleCount => _samples.Count;

        private readonly long _dataOffset;
        public SampledStream(int w, int h, int bitpix, long dataOff) 
        {
            _dataOffset = dataOff;
            _length = dataOff + (long)w * h * (Math.Abs(bitpix)/8);
        }
        public void WriteSample(long offset, byte[] data) {
            byte[] copy = new byte[data.Length]; Array.Copy(data, copy, data.Length);
            _samples[offset] = copy; BufferedBytes += data.Length;
        }
        public override int Read(byte[] buffer, int offset, int count) {
            if (_pos >= _length) return 0;
            int toRead = (int)Math.Min(count, _length - _pos);
            Array.Clear(buffer, offset, toRead);
            
            // Find the nearest sample with the same vertical parity (for Bayer/Color integrity).
            long bestKey = -1;
            long minDiff = long.MaxValue;
            long relPos = _pos - _dataOffset;
            
            if (relPos < 0) {
                // Header area: just find absolute nearest
                foreach (var k in _samples.Keys) {
                    long diff = Math.Abs(k - _pos);
                    if (diff < minDiff) { minDiff = diff; bestKey = k; }
                }
            } else {
                // Data area: match row parity if possible
                long rowLen = (_samples.Count > 1) ? _samples.Values.ElementAt(1).Length : 1; // Element 0 is header
                long requestedRowIdx = relPos / rowLen;

                foreach (var k in _samples.Keys) {
                    if (k < _dataOffset) continue;
                    long sampleRowIdx = (k - _dataOffset) / rowLen;
                    if (sampleRowIdx % 2 != requestedRowIdx % 2) continue; // Match parity
                    
                    long diff = Math.Abs(k - _pos);
                    if (diff < minDiff) { minDiff = diff; bestKey = k; }
                }
                
                // Fallback to absolute nearest if no parity match found
                if (bestKey == -1) {
                    foreach (var k in _samples.Keys) {
                        long diff = Math.Abs(k - _pos);
                        if (diff < minDiff) { minDiff = diff; bestKey = k; }
                    }
                }
            }

            if (bestKey != -1) {
                byte[] sample = _samples[bestKey];
                long rowLen = sample.Length;
                int colIdx = (int)((_pos - _dataOffset) % rowLen);
                if (colIdx < 0) colIdx = (int)(_pos % rowLen); // Header fallback
                
                int n = Math.Min(toRead, (int)rowLen - colIdx);
                Array.Copy(sample, colIdx, buffer, offset, n);
            }
            
            _pos += toRead; return toRead;
        }
        public override long Seek(long offset, SeekOrigin origin) {
            if (origin == SeekOrigin.Begin) _pos = offset;
            else if (origin == SeekOrigin.Current) _pos += offset;
            else _pos = _length + offset;
            return _pos;
        }
        public override void Flush() { }
        public override void SetLength(long value) { }
        public override void Write(byte[] buffer, int offset, int count) { }
    }

    // ── Image metadata extracted from the FITS primary header ───────────
    public struct ImageInfo
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
        public string ImageType;   // IMAGETYP or FRAME keyword (e.g. Light Frame / Dark Frame / Flat Field / Bias Frame)

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

        private static readonly int _pid = System.Diagnostics.Process.GetCurrentProcess().Id;

        internal static void Log(string msg)
        {
            if (!Settings.EnableTracing) return;
            msg = $"[{DateTime.Now:HH:mm:ss.fff}] [PID:{_pid}] [UI] [T{System.Threading.Thread.CurrentThread.ManagedThreadId}] {msg}";
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
        private StatusStrip   _statusStrip;
        private ToolStripStatusLabel _statusLabel;
        private ToolStripStatusLabel _versionLabel;

        // ── Concurrency ─────────────────────────────────────────────────
        private System.Threading.CancellationTokenSource _cts;
	private Task _imageLoadTask;
        private readonly object _syncRoot = new object();

        private void CancelOldLoad()
        {
            lock (_syncRoot)
            {
                if (_cts != null)
                {
                    _cts.Cancel();
                    //_cts.Dispose();
		    // NO dispose aquí, se hará en Dispose tras esperar a la tarea
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

        protected override void Dispose(bool disposing)
        {
            Log($"Dispose({disposing}).BEGIN");
            try
            {
                if (disposing)
                {
                    lock (_syncRoot)
                    {
                        if (_cts != null)
                        {
                            _cts.Cancel();
                            // We don't wait for the task anymore (Fire & Forget).
                            // The task's finally block will handle cleanup.
                            _cts.Dispose();
                            _cts = null;
                        }
                    }
                    
                    if (_pictureBox.Image != null)
                    {
                        try { _pictureBox.Image.Dispose(); } catch {}
                        _pictureBox.Image = null;
                    }
                }
            }
            catch (Exception ex) { Log("Dispose EXCEPTION: " + ex); }
            finally { base.Dispose(disposing); Log("Dispose.END"); }
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
                Text = Strings.Title,
                ForeColor = Color.FromArgb(140, 200, 255),
                Font = new Font("Segoe UI", 10f, FontStyle.Bold)
            };
            _lblLogStatus = new Label
            {
                AutoSize = true, Dock = DockStyle.Right,
                TextAlign = ContentAlignment.MiddleRight,
                Text = Strings.LogStatus(false),
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

            // ── Status bar ──────────────────────────────────────────────
            _statusStrip = new StatusStrip { BackColor = Color.FromArgb(40, 40, 60), SizingGrip = false };
            _statusLabel = new ToolStripStatusLabel { 
                Text = Strings.RightClickHint, 
                ForeColor = Color.FromArgb(160, 160, 180),
                Spring = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            _versionLabel = new ToolStripStatusLabel { 
                Text = $"v1.3.0 · 2026-04-14 20:15", 
                ForeColor = Color.FromArgb(110, 110, 140),
                Font = new Font("Segoe UI", 8f)
            };
            _statusStrip.Items.Add(_statusLabel);
            _statusStrip.Items.Add(_versionLabel);

            ctxMenu.Opening += (s, e) => {
                ctxMenu.Items.Clear();
                bool currentShowImg = Settings.ShowImage;
                bool currentLogOn = Settings.EnableTracing;

                var itemVer = new ToolStripMenuItem($"Fits Preview v1.3.0") { Enabled = false };
                var itemDate = new ToolStripMenuItem($"Built: 2026-04-14 20:15") { Enabled = false };
                ctxMenu.Items.Add(itemVer);
                ctxMenu.Items.Add(itemDate);
                ctxMenu.Items.Add(new ToolStripSeparator());

                var itemImg = new ToolStripMenuItem(currentShowImg ? Strings.MenuHideImage : Strings.MenuShowImage);
                itemImg.Click += (sender, args) => {
                    Settings.ShowImage = !currentShowImg;
                    _statusLabel.Text = Strings.MenuSavedImage;
                    _statusLabel.ForeColor = Color.FromArgb(180, 255, 180);
                };
                
                var itemLog = new ToolStripMenuItem(currentLogOn ? Strings.MenuDisableTrace : Strings.MenuEnableTrace);
                itemLog.Click += (sender, args) => {
                    Settings.EnableTracing = !currentLogOn;
                    _statusLabel.Text = Strings.MenuSavedTrace;
                    _statusLabel.ForeColor = Color.FromArgb(180, 255, 180);
                };

                ctxMenu.Items.Add(itemImg);
                ctxMenu.Items.Add(itemLog);
                
                ctxMenu.Items.Add(new ToolStripSeparator());

                if (_pictureBox.Image != null && currentShowImg)
                {
                    var itemCopyImg = new ToolStripMenuItem(Strings.MenuCopyImage(_pictureBox.Image.Width, _pictureBox.Image.Height));
                    itemCopyImg.Click += (sender, args) => {
                        try { Clipboard.SetImage(_pictureBox.Image); } catch { }
                    };
                    ctxMenu.Items.Add(itemCopyImg);
                }
                
                if (_gridHeader.SelectedRows.Count > 0)
                {
                    var itemCopy = new ToolStripMenuItem(Strings.MenuCopyRow);
                    itemCopy.Click += (sender, args) => {
                        var row = _gridHeader.SelectedRows[0];
                        string text = $"{row.Cells[0].Value}={row.Cells[1].Value} // {row.Cells[2].Value}";
                        try { Clipboard.SetText(text); } catch { }
                    };
                    ctxMenu.Items.Add(itemCopy);
                }

                var itemCopyCsv = new ToolStripMenuItem(Strings.MenuCopyCsv);
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
            Controls.Add(_statusStrip);
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
                    ShowError(Strings.FileNotFound + filePath);
                    return;
                }
                // FileShare.Delete allows the parent folder to be renamed/moved even while
                // we hold this handle open (async image task may still be reading).
                // ownsStream=true hands disposal responsibility to LoadFits/StartImageLoad.
                var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                                        FileShare.ReadWrite | FileShare.Delete);
                LoadFits(fs, Path.GetFileName(filePath), ownsStream: true);
            }
            catch (Exception ex)
            {
                Log("LoadFits(string) — EXCEPTION: " + ex);
                ShowError(Strings.ErrorOpening + ex.Message);
            }
        } 
        public void LoadFits(Stream stream, string fileName, bool ownsStream = false, List<(string, string, string)> preParsedKeywords = null, ImageInfo? preParsedInfo = null)
        {
            Log($"LoadFits.BEGIN — '{fileName}' (owns={ownsStream})");
            try
            {
                bool showImg = Settings.ShowImage;
                bool logOn   = Settings.EnableTracing;

                Action updateLayout = () => {
                    _pictureBox.Visible = showImg;
                    _split.Panel1Collapsed = !showImg;
                    if (!showImg) _lblProgress.Visible = false;
                    _lblImageHint.Text = Strings.RightClickHint;
                    _lblImageHint.Visible = true;
                    _lblLogStatus.Text = Strings.LogStatus(logOn);
                    _lblLogStatus.ForeColor = logOn ? Color.FromArgb(150, 255, 150) : Color.FromArgb(120, 120, 140);
                };

                if (InvokeRequired) Invoke(updateLayout); else updateLayout();

                List<(string, string, string)> rows;
                ImageInfo info;

                if (preParsedKeywords != null && preParsedInfo.HasValue)
                {
                    Log("  Using pre-parsed keywords/info");
                    rows = preParsedKeywords;
                    info = preParsedInfo.Value;
                }
                else
                {
                    Log("  Parsing stream...");
                    (rows, info) = ParseFitsStream(stream);
                }

                Action populate = () => PopulateGrid(rows, fileName, info);
                if (InvokeRequired) Invoke(populate); else populate();

                if (showImg && info.HasImage && !IsDisposed)
                {
                    CancelOldLoad();
                    // Task starts, samples the stream, then releases it immediately
                    StartImageLoad(stream, info, _cts.Token, fileName, ownsStream);
                }
                else
                {
                    if (ownsStream) { try { stream.Dispose(); } catch {} }
                    if (!showImg) Log("  Image loading disabled in registry");
                    else if (!info.HasImage) Log("  No image data found in header");
                }
            }
            catch (Exception ex)
            {
                Log("LoadFits EXCEPTION: " + ex);
                ShowError(Strings.ErrorReadingStream + ex.Message);
                if (ownsStream) try { stream.Dispose(); } catch { }
            }
            finally { Log("LoadFits.END"); }
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
                : Strings.NoImageHint;
            
            if (!Settings.ShowImage) imgNote += Strings.ImageHiddenHint;
            _lblTitle.Text = fileName + imgNote;

            foreach (var (kw, val, comment) in rows)
            {
                int idx = _gridHeader.Rows.Add(kw, val, comment);
                if      (kw == "END")                       _gridHeader.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(255, 120, 120);
                else if (kw == "COMMENT" || kw == "HISTORY") _gridHeader.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(160, 200, 140);
                else if (kw.StartsWith("NAXIS"))             _gridHeader.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(255, 200, 100);
            }
            
            _gridHeader.Update(); // Force table paint immediately
            this.Update(); // Force window paint immediately
            Log("PopulateGrid — done");
        }

        // ── Async image load ────────────────────────────────────────────
        private void StartImageLoad(Stream stream, ImageInfo info, System.Threading.CancellationToken token, string fileName, bool ownsStream = false)
        {
            Log("StartImageLoad.BEGIN");
            _imageLoadTask = Task.Run(() =>
            {
                Log("  Background Task — started");
                try
                {
                    if (token.IsCancellationRequested) { Log("  Background Task — cancelled before start"); if (ownsStream) stream.Dispose(); return; }

                    Stream sourceStream = stream;
                    if (!(stream is SampledStream))
                    {
                        Log("  Background Task — creating SampledStream...");
                        void BufferReport(string msg) => TryBeginInvoke(new Action(() => {
                            if (!IsDisposed && _cts != null && !_cts.IsCancellationRequested) {
                                _lblProgress.Text = msg;
                                _lblProgress.Visible = true;
                            }
                        }));

                        BufferReport(Strings.RenderLoading);
                        int targetDim = 600;
                        var ss = new SampledStream(info.Width, info.Height, info.BitPix, info.DataOffset);
                        
                        sourceStream.Seek(0, SeekOrigin.Begin);
                        byte[] hbuf = new byte[info.DataOffset > 0 ? info.DataOffset : 2880];
                        sourceStream.Read(hbuf, 0, hbuf.Length);
                        ss.WriteSample(0, hbuf);

                        double aspect = (double)info.Width / info.Height;
                        int outW = info.Width >= info.Height ? Math.Min(info.Width, targetDim) : Math.Max(1, (int)(Math.Min(info.Height, targetDim) * aspect));
                        int outH = info.Width >= info.Height ? Math.Max(1, (int)(outW / aspect)) : Math.Min(info.Height, targetDim);
                        bool isBayer = !string.IsNullOrEmpty(info.BayerPattern) && info.Planes == 1;
                        int sy = Math.Max(1, info.Height / outH);
                        if (isBayer) sy = (sy / 2) * 2;
                        if (sy < 1) sy = 1;

                        int bpp = info.BytesPerPixel;
                        long rowBytes = (long)info.Width * bpp;
                        byte[] row = new byte[rowBytes];

                        for (int p = 0; p < info.Planes; p++)
                        {
                            if (token.IsCancellationRequested) break;
                            long planeOff = info.DataOffset + (long)p * info.Height * rowBytes;
                            for (int y = 0; y <= info.Height - (isBayer ? 2 : 1); y += sy)
                            {
                                if (token.IsCancellationRequested) break;
                                if (y % (sy * 20) == 0) BufferReport(Strings.RenderProgress(y * 100 / info.Height));

                                sourceStream.Seek(planeOff + (long)y * rowBytes, SeekOrigin.Begin);
                                sourceStream.Read(row, 0, (int)rowBytes);
                                ss.WriteSample(planeOff + (long)y * rowBytes, row);
                                if (isBayer) {
                                    sourceStream.Read(row, 0, (int)rowBytes);
                                    ss.WriteSample(planeOff + (long)(y+1) * rowBytes, row);
                                }
                            }
                        }
                        
                        if (token.IsCancellationRequested) {
                            Log("  Background Task — cancelled during sampling");
                            if (ownsStream) sourceStream.Dispose();
                            ss.Dispose();
                            return;
                        }

                        ss.Position = 0;
                        if (ownsStream) sourceStream.Dispose();
                        sourceStream = ss;
                        Log($"  Background Task — switched to SampledStream ({ss.BufferedBytes/1024} KB).");
                    }

                    void Report(string msg) => TryBeginInvoke(new Action(() => {
                        if (!IsDisposed && _cts != null && !_cts.IsCancellationRequested) {
                            _lblProgress.Text = msg;
                            _lblProgress.Visible = true;
                        }
                    }));

                    Log("  Background Task — calling RenderImage...");
                    Bitmap bmp = RenderImage(sourceStream, info, token, Report, maxDim: 600);
                    if (bmp != null && !token.IsCancellationRequested && !IsDisposed)
                    {
                        TryBeginInvoke(new Action(() =>
                        {
                            if (_pictureBox.Image != null) { try { _pictureBox.Image.Dispose(); } catch {} _pictureBox.Image = null; }
                            _pictureBox.Image = bmp;
                            _pictureBox.Visible = true;
                            _lblProgress.Visible = false;
                        }));
                        Log("  Background Task — image rendered and posted to UI");
                    }
                    else if (bmp != null) bmp.Dispose();
                    
                    if (sourceStream is SampledStream) sourceStream.Dispose();
                }
                catch (Exception ex) { if (!token.IsCancellationRequested) Log("Background Task EXCEPTION: " + ex); }
                finally 
                { 
                    Log("  Background Task — finished"); 
                    // Now that we have the full image, let Explorer update the thumbnail
                    NotifyFileChanged(fileName);
                }
            });
            Log("StartImageLoad.END");
        }

        private void TryBeginInvoke(Action action)
        {
            try { if (!IsDisposed && IsHandleCreated) BeginInvoke(action); } catch {}
        }

public void WaitRenderTask(int timeoutMs = 2000)
{
    lock (_syncRoot)
    {
        if (_cts != null) { _cts.Cancel(); _cts.Dispose(); _cts = null; }
    }
    // Esperamos brevemente a que el ThreadPool termine el finally.
    // Timeout corto para evitar deadlocks si la tarea está bloqueada en I/O de nube.
    _imageLoadTask?.Wait(timeoutMs);
}


        // ── Rendering pipeline ──────────────────────────────────────────
        private static Bitmap RenderImage(Stream stream, ImageInfo info, System.Threading.CancellationToken token, Action<string> reportProgress = null, int maxDim = 2000, bool disposeSourceAfterRead = false)
        {
            try {
                int MAX_DIM = maxDim; 
                double aspect = (double)info.Width / info.Height;
                int outW, outH;
                if (info.Width >= info.Height)
                { outW = Math.Min(info.Width, MAX_DIM); outH = Math.Max(1, (int)(outW / aspect)); }
                else
                { outH = Math.Min(info.Height, MAX_DIM); outW = Math.Max(1, (int)(outH * aspect)); }

                Log($"RenderImage.BEGIN — target {outW}×{outH}");

                // 1. READ phase (IO intensive, holds the lock)
                float[][] planes = ReadAllPlanes(stream, info, outW, outH, out int aW, out int aH, token, reportProgress);
                
                // 2. CRITICAL: Release the file lock as soon as bytes are in memory
                if (disposeSourceAfterRead) {
                    Log("  RenderImage — Read done, releasing source stream handle...");
                    try { stream.Dispose(); } catch { }
                }

                if (token.IsCancellationRequested || planes == null) return null;
                Log($"RenderImage — planes={planes.Length} actual={aW}×{aH}");

                // 3. PROCESS phase (CPU intensive, file is already UNLOCKED)
                if (planes.Length >= 3)
                    return PlanesToColorBitmap(planes[0], planes[1], planes[2], aW, aH, token);

                AdaptiveStretch(planes[0], out float low, out float high);
                if (token.IsCancellationRequested) return null;
                
                Color tint = GetFilterTint(info.Filter);
                return PixelsToBitmap(planes[0], aW, aH, low, high, tint);
            }
            finally {
                if (disposeSourceAfterRead) { try { stream.Dispose(); } catch { } }
            }
        }

        private static float[][] ReadAllPlanes(
            Stream stream, ImageInfo info,
            int targetW, int targetH,
            out int actualW, out int actualH,
            System.Threading.CancellationToken token,
            Action<string> reportProgress = null)
        {
            try
            {
                bool isBayer = !string.IsNullOrEmpty(info.BayerPattern) && info.Planes == 1;
                int sx = Math.Max(1, info.Width / targetW);
                int sy = Math.Max(1, info.Height / targetH);

                if (isBayer) {
                    sx = (sx < 2) ? 2 : (sx % 2 != 0 ? sx + 1 : sx);
                    sy = (sy < 2) ? 2 : (sy % 2 != 0 ? sy + 1 : sy);
                }

                actualW = 0; for (int x = 0; x <= info.Width - (isBayer ? 2 : 1); x += sx) actualW++;
                actualH = 0; for (int y = 0; y <= info.Height - (isBayer ? 2 : 1); y += sy) actualH++;

                int bpp = info.BytesPerPixel;
                long rowBytes = (long)info.Width * bpp;
                int outPlanes = isBayer ? 3 : info.Planes;

                float[][] planes = new float[outPlanes][];
                for (int p = 0; p < outPlanes; p++) planes[p] = new float[actualW * actualH];

                for (int p = 0; p < info.Planes && p < (isBayer ? 1 : outPlanes); p++)
                {
                    long planeOff = info.DataOffset + (long)p * info.Height * rowBytes;
                    long lastY = (long)(actualH - 1) * sy + (isBayer ? 1 : 0);
                    if (lastY >= info.Height) lastY = info.Height - 1;
                    
                    long start = planeOff;
                    long end = planeOff + (lastY * rowBytes) + rowBytes;
                    long totalRange = end - start;

                    if (totalRange > 0 && totalRange < 32 * 1024 * 1024)
                    {
                        Log($"  ReadAllPlanes — Block read {totalRange / 1024} KB...");
                        byte[] block = new byte[totalRange];
                        stream.Seek(start, SeekOrigin.Begin);
                        int read = 0; while(read < totalRange) {
                            int r = stream.Read(block, read, (int)(totalRange - read));
                            if (r <= 0) break; read += r;
                        }

                        int outY = 0;
                        for (int y = 0; y <= info.Height - (isBayer ? 2 : 1) && outY < actualH; y += sy)
                        {
                            if (token.IsCancellationRequested) return null;
                            int rowOff = (int)((long)y * rowBytes);
                            FillRow(planes, outY, actualW, block, rowOff, (int)(rowOff + rowBytes), info, sx, actualH, isBayer, p);
                            outY++;
                        }
                    }
                    else
                    {
                        Log($"  ReadAllPlanes — Range {totalRange/1024/1024}MB too large, using row-skipping.");
                        byte[] rowBuf = new byte[rowBytes];
                        byte[] nextRowBuf = isBayer ? new byte[rowBytes] : null;
                        int outY = 0;
                        for (int y = 0; y <= info.Height - (isBayer ? 2 : 1) && outY < actualH; y += sy)
                        {
                            if (token.IsCancellationRequested) break;
                            stream.Seek(planeOff + (long)y * rowBytes, SeekOrigin.Begin);
                            stream.Read(rowBuf, 0, (int)rowBytes);
                            if (isBayer) {
                                stream.Seek(planeOff + (long)(y+1) * rowBytes, SeekOrigin.Begin);
                                stream.Read(nextRowBuf, 0, (int)rowBytes);
                            }
                            
                            FillRow(planes, outY, actualW, rowBuf, 0, 0, info, sx, actualH, isBayer, p, nextRowBuf);
                            outY++;
                        }
                    }
                }
                return planes;
            }
            catch (Exception ex) { 
                Log("ReadAllPlanes (Block) ERROR: " + ex); 
                actualW = 0; actualH = 0;
                return null; 
            }
        }

        private static void FillRow(float[][] planes, int outY, int outW, byte[] buffer, int offset, int nextRowOffset, ImageInfo info, int sx, int actualH, bool isBayer, int p, byte[] nextRowBuffer = null)
        {
            int bpp = info.BytesPerPixel;
            int outIdxBase = (actualH - 1 - outY) * outW;
            string pat = (info.BayerPattern ?? "RGGB").ToUpper();
            
            for (int x = 0; x < outW; x++)
            {
                int targetX = x * sx;
                int pxOff = offset + (targetX * bpp);
                
                if (isBayer)
                {
                    double p00 = ReadRaw(buffer, pxOff, info.BitPix);
                    double p10 = ReadRaw(buffer, pxOff + bpp, info.BitPix);
                    
                    byte[] b2 = nextRowBuffer ?? buffer;
                    int off2 = nextRowBuffer != null ? (targetX * bpp) : nextRowOffset + (targetX * bpp);
                    
                    double p01 = ReadRaw(b2, off2, info.BitPix);
                    double p11 = ReadRaw(b2, off2 + bpp, info.BitPix);
                    
                    float r, g, b;
                    if (pat == "RGGB") { r=(float)p00; g=(float)(p10+p01)/2f; b=(float)p11; }
                    else if (pat == "GRBG") { g=(float)(p00+p11)/2f; r=(float)p10; b=(float)p01; }
                    else if (pat == "GBRG") { g=(float)(p00+p11)/2f; b=(float)p10; r=(float)p01; }
                    else { b=(float)p00; g=(float)(p10+p01)/2f; r=(float)p11; }

                    planes[0][outIdxBase + x] = (float)(r * info.BScale + info.BZero);
                    planes[1][outIdxBase + x] = (float)(g * info.BScale + info.BZero);
                    planes[2][outIdxBase + x] = (float)(b * info.BScale + info.BZero);
                }
                else
                {
                    double raw = ReadRaw(buffer, pxOff, info.BitPix);
                    planes[p][outIdxBase + x] = (float)(raw * info.BScale + info.BZero);
                }
            }
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

            // 1. Fast Bail for non-FITS: first 8 bytes must be 'SIMPLE  '
            byte[] magic = new byte[8];
            stream.Read(magic, 0, 8);
            if (Encoding.ASCII.GetString(magic) != "SIMPLE  ") {
                Log("ParseFitsStream — Fast-Bail: not a FITS (no 'SIMPLE  ' keyword)");
                return (result, info);
            }
            stream.Seek(0, SeekOrigin.Begin);

            using (var br = new BinaryReader(stream, Encoding.ASCII, true)) // leaveOpen=true
            {
                bool endFound = false;
                int  blockIdx = 0;
                const int blockLimit = 256; // Cap at ~737KB to prevent hangs on invalid files

                while (!endFound && blockIdx < blockLimit)
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
                                case "IMAGETYP":
                                case "FRAME":
                                    if (info.ImageType == null) info.ImageType = val.Trim(); break;
                            }
                        }
                        else
                        {
                            result.Add((kw, "", rest.Trim()));
                        }
                    }
                    blockIdx++;
                }

                if (!endFound)
                {
                    if (blockIdx >= blockLimit) Log($"ParseFitsStream — WARNING: Reached limit ({blockLimit}) without END.");
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

        // ── Thumbnail & Badge rendering (called by IThumbnailProvider) ─────────

        /// <summary>
        /// Renders a stride-sampled thumbnail fitting within <paramref name="cx"/>×<paramref name="cx"/> pixels.
        /// For a 4656-wide sensor at cx=256 the stride is ~18, so only ≈1/324th of the pixel
        /// data is read — the rest of the file is never touched.
        /// </summary>
        internal static Bitmap RenderThumbnail(Stream stream, ImageInfo info, int cx, bool ownsStream = false)
            => RenderImage(stream, info, System.Threading.CancellationToken.None, null, cx, disposeSourceAfterRead: ownsStream);

                /// <summary>
        /// Renders a static identification badge when image rendering is disabled (ShowImage=false).
        /// Encodes the frame type (LIGHT/DARK/FLAT/BIAS) via background colour and icon,
        /// and the FILTER keyword via an accent colour + text label.
        /// Zero pixel data is read — uses only the already-parsed <see cref="ImageInfo"/> metadata.
        /// </summary>

    ///  <summary >
    /// Renders a "Top-Heavy" identification badge with dynamic font fitting.
    /// - Positions text in the UPPER half to avoid folder overlays.
    /// - Uses a loop to shrink font size so it NEVER gets cut off.
    /// - Improved "FITS" subtitle visibility (Solid color, larger, bold).
    ///  </summary >
    internal static Bitmap RenderStaticBadge(ImageInfo info, int size)
    {
        string typeRaw = (info.ImageType ?? string.Empty).ToLowerInvariant().Trim();

        string frameLabel = "FITS";
        Color bgColor;
        Color textColor;

        // Paleta de colores
        if      (typeRaw.Contains("light")) { frameLabel = "LIGHT"; bgColor = Color.FromArgb(12, 16, 35); textColor = Color.FromArgb(220, 230, 255); }
        else if (typeRaw.Contains("flat"))  { frameLabel = "FLAT";  bgColor = Color.FromArgb(215, 220, 230); textColor = Color.FromArgb(20, 25, 35); }
        else if (typeRaw.Contains("dark"))  { frameLabel = "DARK";  bgColor = Color.FromArgb(40, 10, 10); textColor = Color.FromArgb(255, 180, 180); }
        else if (typeRaw.Contains("bias"))  { frameLabel = "BIAS";  bgColor = Color.FromArgb(18, 18, 25); textColor = Color.FromArgb(190, 190, 210); }
        else                                { frameLabel = "FITS";  bgColor = Color.FromArgb(20, 20, 30); textColor = Color.Silver; }

        Color accentColor = GetFilterAccentColor(info.Filter);

        var bmp = new Bitmap(size, size, PixelFormat.Format24bppRgb);
        try
        {
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                g.Clear(bgColor);

                // 1. BARRA SUPERIOR (Filtro)
                int stripeH = Math.Max(2, size / 12);
                using (var sb = new SolidBrush(accentColor))
                    g.FillRectangle(sb, 0, 0, size, stripeH);

                // 2. ZONA SEGURA (Zona "Top-Heavy")
                // Definimos un rectángulo en la parte SUPERIOR donde el texto DEBE caber.
                float padding = size * 0.05f;
                RectangleF textRect = new RectangleF(
                    padding,
                    stripeH + padding,
                    size - (padding * 2),
                    (size * 0.55f) - stripeH // Ocupa el 55% superior disponible
                );

                // 3. BUCLE DE AJUSTE DE FUENTE (Font Fitting)
                // Empezamos con un tamaño grande y lo reducimos hasta que mida menos que el ancho del rect.
                float fontSize = size * 0.50f;
                SizeF textSize = SizeF.Empty;

                using (var testBrush = new SolidBrush(textColor))
                {
                    while (fontSize > 6f) // Mínimo legible 6px
                    {
                        using (var testFont = new Font("Segoe UI", fontSize, FontStyle.Bold, GraphicsUnit.Pixel))
                        {
                            textSize = g.MeasureString(frameLabel, testFont);
                            if (textSize.Width <= textRect.Width)
                                break; // ¡Cabe perfectamente!

                            fontSize -= 1f; // Reducir 1 píxel y probar de nuevo
                        }
                    }

                    // 4. DIBUJAR TEXTO PRINCIPAL
                    using (var mainFont = new Font("Segoe UI", fontSize, FontStyle.Bold, GraphicsUnit.Pixel))
                    {
                        using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Near })
                        {
                            g.DrawString(frameLabel, mainFont, testBrush, textRect, sf);
                        }
                    }
                }

                // 5. ETIQUETA "FITS" (Pie de página)
                // MEJORADA: Más grande, negrita y color sólido de alto contraste
                float footerFontSize = size * 0.16f;
                if (footerFontSize < 7f) footerFontSize = 7f; // Mínimo 7px

                // Color sólido inteligente: Oscuro sobre fondo claro, Claro sobre fondo oscuro
                Color footerColor = bgColor.GetBrightness() > 0.5f
                    ? Color.FromArgb(80, 80, 90)   // Gris oscuro para fondos claros (FLAT)
                    : Color.FromArgb(220, 220, 230); // Blanco hueso para fondos oscuros

                using (var footerFont = new Font("Segoe UI", footerFontSize, FontStyle.Bold, GraphicsUnit.Pixel))
                using (var footerBrush = new SolidBrush(footerColor))
                {
                    string footerText = "FITS";
                    SizeF fSize = g.MeasureString(footerText, footerFont);

                    if (fSize.Width <= size * 0.95f)
                    {
                        PointF fPt = new PointF(
                            (size - fSize.Width) / 2f,
                            size - fSize.Height - (size * 0.05f) // Margen inferior
                        );
                        g.DrawString(footerText, footerFont, footerBrush, fPt);
                    }
                }
            }
            return bmp;
        }
        catch
        {
            bmp?.Dispose();
            return null;
        }
    }
    
        
        /// <summary>Maps a FITS FILTER keyword value to a badge accent colour.</summary>
        private static Color GetFilterAccentColor(string filter)
        {
            if (string.IsNullOrEmpty(filter)) return Color.FromArgb(136, 136, 187);
            string f = filter.ToLowerInvariant().Trim();

            // Narrowband emission lines — check most specific tokens first
            if (f == "ha" || f == "h-a" || f == "h\u03b1" ||
                f.StartsWith("halpha") || f.Contains("h-alpha") || f.Contains("h_alpha"))
                return Color.FromArgb(255,  60,  60);
            if (f.Contains("oiii") || f == "o3" || f.Contains("o-iii") || f.Contains("o_iii"))
                return Color.FromArgb( 60, 240, 240);
            if (f.Contains("sii")  || f == "s2" || f.Contains("s-ii")  || f.Contains("s_ii"))
                return Color.FromArgb(255, 128,  32);
            if (f == "hb" || f == "h-b" || f == "h\u03b2" ||
                f.StartsWith("hbeta") || f.Contains("h-beta") || f.Contains("h_beta"))
                return Color.FromArgb( 80,  80, 255);

            // Broadband LRGB
            if (f == "l" || f == "lum" || f == "luminance" || f.StartsWith("lum"))
                return Color.FromArgb(220, 220, 220);
            if (f == "r" || f == "red")
                return Color.FromArgb(255,  50,  50);
            if (f == "g" || f == "green" || f == "grn")
                return Color.FromArgb( 60, 220,  60);
            if (f == "b" || f == "blue" || f == "blu")
                return Color.FromArgb( 60, 100, 255);

            // Unknown filter → neutral purple-grey
            return Color.FromArgb(136, 136, 187);
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

        [DllImport("shell32.dll")]
        private static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);
        
        public static void NotifyFileChanged(string path)
        {
            if (string.IsNullOrEmpty(path) || path == "FITS Stream") return;
            try 
            {
                IntPtr pidl = IntPtr.Zero;
                // SHCNE_UPDATEITEM = 0x00002000, SHCNF_PATHW = 0x0005
                IntPtr pathPtr = Marshal.StringToCoTaskMemUni(path);
                SHChangeNotify(0x2000, 0x0005, pathPtr, IntPtr.Zero);
                Marshal.FreeCoTaskMem(pathPtr);
            }
            catch { }
        }
    }
}
