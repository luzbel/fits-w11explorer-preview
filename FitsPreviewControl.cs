using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace FitsPreviewHandler
{
    public class FitsPreviewControl : UserControl
    {
        // ── Shared log path (same folder as the DLL so it's always findable) ──
        internal static readonly string LogPath =
            Path.Combine(
                Path.GetDirectoryName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location) ?? @"C:\Temp",
                "fits_trace.log");

        internal static void Log(string msg)
        {
            try
            {
                File.AppendAllText(LogPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [Control] [T{System.Threading.Thread.CurrentThread.ManagedThreadId}] {msg}\n");
            }
            catch { /* never crash because of logging */ }
        }

        // ────────────────────────────────────────────────────────────────────
        private DataGridView _gridHeader;
        private Label _lblTitle;
        private Panel _topPanel;
        private TextBox _txtError;

        public FitsPreviewControl()
        {
            Log("FitsPreviewControl.ctor — start");
            try
            {
                InitializeComponent();
                Log("FitsPreviewControl.ctor — InitializeComponent OK");
            }
            catch (Exception ex)
            {
                Log("FitsPreviewControl.ctor — EXCEPTION: " + ex);
                throw;
            }
        }

        private void InitializeComponent()
        {
            Log("InitializeComponent — begin");

            this.BackColor = Color.FromArgb(24, 24, 36);
            this.ForeColor = Color.FromArgb(220, 220, 235);

            // ── Top title bar ─────────────────────────────────────────
            _topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 36,
                BackColor = Color.FromArgb(40, 40, 60),
                Padding = new Padding(8, 0, 0, 0)
            };

            _lblTitle = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "FITS Header",
                ForeColor = Color.FromArgb(140, 200, 255),
                Font = new Font("Segoe UI", 10f, FontStyle.Bold)
            };
            _topPanel.Controls.Add(_lblTitle);
            Log("InitializeComponent — topPanel created");

            // ── Header grid ───────────────────────────────────────────
            _gridHeader = new DataGridView
            {
                Dock = DockStyle.Fill,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                BackgroundColor = Color.FromArgb(28, 28, 42),
                GridColor = Color.FromArgb(50, 50, 70),
                BorderStyle = BorderStyle.None,
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
                ScrollBars = ScrollBars.Both,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            };

            // Columns: Keyword | Value | Comment
            _gridHeader.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Keyword",
                HeaderText = "Keyword",
                Width = 90,
                MinimumWidth = 60,
                Resizable = DataGridViewTriState.True
            });
            _gridHeader.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Value",
                HeaderText = "Value",
                Width = 160,
                MinimumWidth = 80,
                Resizable = DataGridViewTriState.True
            });
            _gridHeader.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Comment",
                HeaderText = "Comment",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 80,
                Resizable = DataGridViewTriState.True
            });

            // Alternate row color
            _gridHeader.RowsAdded += (s, e) =>
            {
                for (int i = e.RowIndex; i < e.RowIndex + e.RowCount && i < _gridHeader.Rows.Count; i++)
                {
                    _gridHeader.Rows[i].DefaultCellStyle.BackColor =
                        (i % 2 == 0) ? Color.FromArgb(28, 28, 42) : Color.FromArgb(33, 33, 50);
                }
            };
            Log("InitializeComponent — grid created with 3 columns");

            // ── Error box (hidden until needed) ──────────────────────
            _txtError = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                Visible = false,
                BackColor = Color.FromArgb(60, 20, 20),
                ForeColor = Color.FromArgb(255, 160, 160),
                Font = new Font("Consolas", 9f),
                ScrollBars = ScrollBars.Vertical
            };

            this.Controls.Add(_gridHeader);
            this.Controls.Add(_topPanel);
            this.Controls.Add(_txtError);

            Log("InitializeComponent — controls added to UserControl");
        }

        // ─────────────────────────────────────────────────────────────────
        //  FITS loader (called from the UI thread)
        // ─────────────────────────────────────────────────────────────────
        public void LoadFits(string filePath)
        {
            Log($"LoadFits — filePath='{filePath}'");
            try
            {
                if (!File.Exists(filePath))
                {
                    Log("LoadFits — file NOT FOUND");
                    ShowError($"Archivo no encontrado:\n{filePath}");
                    return;
                }

                long fileSize = new FileInfo(filePath).Length;
                Log($"LoadFits — file exists, size={fileSize} bytes");

                var rows = ParseFitsHeader(filePath);
                Log($"LoadFits — parsed {rows.Count} keyword rows");

                if (this.InvokeRequired)
                {
                    Log("LoadFits — InvokeRequired=true, dispatching to UI thread");
                    this.Invoke(new Action(() => PopulateGrid(rows, filePath)));
                }
                else
                {
                    PopulateGrid(rows, filePath);
                }

                Log("LoadFits — done");
            }
            catch (Exception ex)
            {
                Log("LoadFits — EXCEPTION: " + ex);
                ShowError("Error leyendo archivo FITS:\r\n" + ex.ToString());
            }
        }

        private void PopulateGrid(List<(string kw, string val, string comment)> rows, string filePath)
        {
            Log($"PopulateGrid — {rows.Count} rows, file='{Path.GetFileName(filePath)}'");
            try
            {
                _gridHeader.Rows.Clear();
                _lblTitle.Text = "FITS Header — " + Path.GetFileName(filePath) + $"  ({rows.Count} keywords)";

                foreach (var (kw, val, comment) in rows)
                {
                    int idx = _gridHeader.Rows.Add(kw, val, comment);

                    // Color-code special keywords
                    if (kw == "END")
                        _gridHeader.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(255, 120, 120);
                    else if (kw == "COMMENT" || kw == "HISTORY")
                        _gridHeader.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(160, 200, 140);
                    else if (kw.StartsWith("NAXIS"))
                        _gridHeader.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(255, 200, 100);
                }

                Log("PopulateGrid — grid filled OK");
            }
            catch (Exception ex)
            {
                Log("PopulateGrid — EXCEPTION: " + ex);
                throw;
            }
        }

        // ─────────────────────────────────────────────────────────────────
        //  FITS header parser
        // ─────────────────────────────────────────────────────────────────
        /// <summary>
        /// Parses the FITS primary header.
        /// FITS spec: header consists of 2880-byte blocks, each containing
        /// 36 records of exactly 80 ASCII bytes. Ends at the END keyword.
        /// </summary>
        private static List<(string kw, string val, string comment)> ParseFitsHeader(string filePath)
        {
            const int RECORD_SIZE = 80;
            const int BLOCK_SIZE = 2880;   // 36 records × 80 bytes

            Log($"ParseFitsHeader — opening '{filePath}'");
            var result = new List<(string, string, string)>();
            int blockIndex = 0;

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new BinaryReader(fs, Encoding.ASCII))
            {
                bool endFound = false;
                while (!endFound && fs.Position < fs.Length)
                {
                    byte[] block = reader.ReadBytes(BLOCK_SIZE);
                    Log($"ParseFitsHeader — block #{blockIndex}, read {block.Length} bytes (pos={fs.Position})");

                    if (block.Length < RECORD_SIZE)
                    {
                        Log("ParseFitsHeader — block too small, stopping");
                        break;
                    }

                    int records = block.Length / RECORD_SIZE;
                    for (int r = 0; r < records && !endFound; r++)
                    {
                        string record = Encoding.ASCII.GetString(block, r * RECORD_SIZE, RECORD_SIZE);
                        string keyword = record.Substring(0, 8).TrimEnd();

                        if (keyword == "END")
                        {
                            Log($"ParseFitsHeader — END found at block #{blockIndex}, record {r}");
                            result.Add(("END", "", ""));
                            endFound = true;
                            break;
                        }

                        if (string.IsNullOrWhiteSpace(keyword))
                            continue;

                        string rawRest = record.Substring(8); // 72 chars

                        if (rawRest.Length >= 2 && rawRest[0] == '=' && rawRest[1] == ' ')
                        {
                            string valueArea = rawRest.Substring(2).TrimStart();
                            ParseValueComment(valueArea, out string val, out string comment);
                            result.Add((keyword, val, comment.Trim()));
                            Log($"ParseFitsHeader   KW='{keyword}' VAL='{val}' CMT='{comment.Trim()}'");
                        }
                        else
                        {
                            string text = rawRest.Trim();
                            result.Add((keyword, "", text));
                            Log($"ParseFitsHeader   KW='{keyword}' (no value) TEXT='{text.Substring(0, Math.Min(40, text.Length))}'");
                        }
                    }
                    blockIndex++;
                }
            }

            Log($"ParseFitsHeader — total keywords: {result.Count}");
            return result;
        }

        /// <summary>
        /// Splits the value+comment area of a FITS keyword record (72 chars after "= ").
        /// Handles string values (quoted with single quotes, internal '' = escaped quote),
        /// and numeric / logical values.
        /// </summary>
        private static void ParseValueComment(string area, out string value, out string comment)
        {
            value = "";
            comment = "";

            if (area.Length == 0) return;

            if (area[0] == '\'')
            {
                // String value: find the closing quote, respecting '' escapes
                int i = 1;
                var sb = new StringBuilder();
                while (i < area.Length)
                {
                    if (area[i] == '\'')
                    {
                        if (i + 1 < area.Length && area[i + 1] == '\'')
                        {
                            sb.Append('\'');
                            i += 2;
                        }
                        else
                        {
                            i++; // skip closing quote
                            break;
                        }
                    }
                    else
                    {
                        sb.Append(area[i]);
                        i++;
                    }
                }
                value = sb.ToString().TrimEnd();

                // Remainder after closing quote may have / comment
                string rest = (i < area.Length) ? area.Substring(i) : "";
                int slash = rest.IndexOf('/');
                comment = slash >= 0 ? rest.Substring(slash + 1).Trim() : "";
            }
            else
            {
                // Numeric or logical value — slash before any string literal is safe
                int slash = area.IndexOf('/');
                if (slash >= 0)
                {
                    value = area.Substring(0, slash).Trim();
                    comment = area.Substring(slash + 1).Trim();
                }
                else
                {
                    value = area.Trim();
                }
            }
        }

        // ─────────────────────────────────────────────────────────────────
        public void ShowError(string message)
        {
            Log("ShowError — " + message.Replace("\r\n", " | ").Replace("\n", " | "));

            void Show()
            {
                _gridHeader.Visible = false;
                _txtError.Visible = true;
                _txtError.Text = message;
            }

            if (this.InvokeRequired)
                this.BeginInvoke(new Action(Show));
            else
                Show();
        }
    }
}
