using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace FitsPreviewHandler
{
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("b7d14566-0509-4cce-a71f-0a554233bdc6")]
    public interface IInitializeWithFile { void Initialize([MarshalAs(UnmanagedType.LPWStr)] string pszFilePath, uint grfMode); }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f")]
    public interface IInitializeWithStream { void Initialize([In, MarshalAs(UnmanagedType.IUnknown)] object pstream, uint grfMode); }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    public interface IShellItem
    {
        void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, out IntPtr ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("7f73be3f-fb79-493c-a6c7-7ee14e245841")]
    public interface IInitializeWithItem { void Initialize([In, MarshalAs(UnmanagedType.IUnknown)] object psi, uint grfMode); }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int left, top, right, bottom; }

    [StructLayout(LayoutKind.Sequential)]
    public struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public Point pt; }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("8895b1c6-b41f-4c1c-a562-0d564250836f")]
    public interface IPreviewHandler
    {
        void SetWindow(IntPtr hwnd, [In] ref RECT rect);
        void SetRect([In] ref RECT rect);
        void DoPreview();
        void Unload();
        void SetFocus();
        void QueryFocus(out IntPtr phwnd);
        [PreserveSig] uint TranslateAccelerator(ref MSG pmsg);
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99")]
    public interface IPropertyStore
    {
        [PreserveSig] uint GetCount(out uint cProps);
        [PreserveSig] uint GetAt(uint iProp, out PropertyKey pkey);
        [PreserveSig] uint GetValue(ref PropertyKey key, out PropVariant pv);
        [PreserveSig] uint SetValue(ref PropertyKey key, ref PropVariant pv);
        [PreserveSig] uint Commit();
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PropertyKey
    {
        public Guid fmtid;
        public uint pid;
        public override bool Equals(object obj) => obj is PropertyKey pk && pk.fmtid == fmtid && pk.pid == pid;
        public override int GetHashCode() => fmtid.GetHashCode() ^ (int)pid;
    }

    [StructLayout(LayoutKind.Explicit, Size = 24)]
    public struct PropVariant : IDisposable
    {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr ptr;
        [FieldOffset(8)] public long value;
        [FieldOffset(8)] public int iVal;
        [FieldOffset(8)] public double dblVal;

        public void SetString(string val) { vt = 31; ptr = Marshal.StringToCoTaskMemUni(val); }
        public void SetInt(int val)       { vt = 3;  iVal = val; }
        public void SetDouble(double val) { vt = 5;  dblVal = val; }

        [DllImport("ole32.dll")]
        public static extern int PropVariantClear(ref PropVariant pvar);
        public void Dispose() { PropVariantClear(ref this); }
    }

    public static class PKEYs
    {
        public static readonly Guid PSG_SUMMARY     = new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
        public static readonly Guid PSG_IMAGE       = new Guid("6444048F-4C8B-11D1-8B70-080036B11A03");
        public static readonly Guid PKEY_Photo      = new Guid("14B81DA1-0135-4D31-96D9-6CBFC9671A99");
        public static readonly Guid PSG_DOC_SUMMARY = new Guid("D5CDD502-2E9C-101B-9397-08002B2CF9AE");

        public static PropertyKey Subject           = new PropertyKey { fmtid = PSG_SUMMARY,     pid = 3 };
        public static PropertyKey Image_Width       = new PropertyKey { fmtid = PSG_IMAGE,       pid = 3 };
        public static PropertyKey Image_Height      = new PropertyKey { fmtid = PSG_IMAGE,       pid = 4 };
        public static PropertyKey Image_BitDepth    = new PropertyKey { fmtid = PSG_IMAGE,       pid = 7 };
        public static PropertyKey Photo_CameraModel = new PropertyKey { fmtid = PKEY_Photo,      pid = 272 };
        public static PropertyKey Photo_Exposure    = new PropertyKey { fmtid = PKEY_Photo,      pid = 33434 };
        public static PropertyKey Category          = new PropertyKey { fmtid = PSG_DOC_SUMMARY, pid = 2 };
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("fc4801a3-2ba9-11cf-a229-00aa003d7352")]
    public interface IObjectWithSite
    {
        void SetSite([In, MarshalAs(UnmanagedType.IUnknown)] object pUnkSite);
        void GetSite(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvSite);
    }

    public enum WTS_ALPHATYPE : int { WTSAT_UNKNOWN = 0, WTSAT_RGB = 1, WTSAT_ARGB = 2 }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("e357fccd-a995-4576-b01f-234630154e96")]
    public interface IThumbnailProvider
    {
        void GetThumbnail(uint cx, out IntPtr hbmp, out WTS_ALPHATYPE pdwAlpha);
    }

    [ComVisible(true)]
    [Guid("AF1C3D6A-81E9-4F5B-9A8C-2D9E71F04B3E")]
    [ClassInterface(ClassInterfaceType.None)]
    public class FitsPreviewHandlerExtension
        : IPreviewHandler, IInitializeWithFile, IInitializeWithStream, IInitializeWithItem,
          IObjectWithSite, IPropertyStore, IThumbnailProvider
    {
        // ── State ───────────────────────────────────────────────────────
        private string _filePath;
        private IntPtr _parentHwnd;
        private RECT   _bounds;
        private FitsPreviewControl           _control;
        private System.Threading.Thread      _uiThread;

        // Metadata is populated eagerly at Initialize time (header-only read, a few KB).
        private ImageInfo? _metadata;
        private List<(string, string, string)> _keywords;

        // COM IStream stored if InitializeWithStream is used.
        private object _comStream;
        private bool   _isSlowLink;

        private static readonly int _pid = System.Diagnostics.Process.GetCurrentProcess().Id;

        // ── Logging ─────────────────────────────────────────────────────
        private void Log(string msg, bool force = false)
        {
            if (!force && !Settings.EnableTracing) return;
            string inst = $"[Inst:{this.GetHashCode():X8}]";
            msg = $"[{DateTime.Now:HH:mm:ss.fff}] [PID:{_pid}] [Ext] [T{System.Threading.Thread.CurrentThread.ManagedThreadId}] {inst} {msg}";
            System.Diagnostics.Debug.WriteLine(msg);
            try { File.AppendAllText(FitsPreviewControl.LogPath, msg + "\n"); } catch { }
        }

        private static void LogStatic(string msg)
        {
            if (!Settings.EnableTracing) return;
            msg = $"[{DateTime.Now:HH:mm:ss.fff}] [PID:{_pid}] [Ext] [T{System.Threading.Thread.CurrentThread.ManagedThreadId}] [Static] {msg}";
            System.Diagnostics.Debug.WriteLine(msg);
            try { File.AppendAllText(FitsPreviewControl.LogPath, msg + "\n"); } catch { }
        }

        static FitsPreviewHandlerExtension()
        {
            LogStatic("--- CLASS LOADED ---");
            AppDomain.CurrentDomain.UnhandledException +=
                (s, e) => LogStatic("!!! UNHANDLED: " + e.ExceptionObject);
        }

        public FitsPreviewHandlerExtension() => Log("=== constructor ===");

        ~FitsPreviewHandlerExtension()
        {
            Log("=== destructor ===");
            try { ReleaseStream(); } catch { }
        }

        private void ReleaseStream()
        {
            if (_comStream != null)
            {
                Log($"  Releasing COM stream {(_comStream?.GetHashCode() ?? 0):X8}...");
                try { Marshal.ReleaseComObject(_comStream); } catch { }
                _comStream = null;
            }
        }

        // ── ParseHeaderFromFile ─────────────────────────────────────────
        private static ImageInfo? ParseHeaderFromFile(string path)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read,
                                               FileShare.ReadWrite | FileShare.Delete))
                {
                    var (_, info) = FitsPreviewControl.ParseFitsStream(fs);
                    return info;
                }
            }
            catch (Exception ex) { LogStatic($"ParseHeaderFromFile failed: {ex.Message}"); return null; }
        }

        // ── IInitializeWithFile ─────────────────────────────────────────
        public void Initialize(string pszFilePath, uint grfMode)
        {
            Log($"IInitializeWithFile.BEGIN — '{pszFilePath}' (mode={grfMode})");
            try
            {
                _filePath    = pszFilePath;
                _comStream   = null;
                _keywords    = null;
                _metadata    = null;

                // Header-only parse (fast, closed immediately)
                using (var fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                {
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    var (rows, info) = FitsPreviewControl.ParseFitsStream(fs);
                    sw.Stop();
                    _keywords = rows;
                    _metadata = info;
                    _isSlowLink = sw.ElapsedMilliseconds > 200;
                    Log($"  Header parsed OK: {info.Width}x{info.Height} (Lat={sw.ElapsedMilliseconds}ms, Slow={_isSlowLink})");
                }
            }
            catch (Exception ex) { Log($"IInitializeWithFile ERROR: {ex.Message}"); }
            finally { Log("IInitializeWithFile.END"); }
        }

        // ── IInitializeWithStream ───────────────────────────────────────
        void IInitializeWithStream.Initialize(object pstream, uint grfMode)
        {
            Log($"IInitializeWithStream.BEGIN — grfMode={grfMode} pstream={pstream?.GetHashCode():X8}");
            try
            {
                // We must hold the stream IF we are in a process that will call DoPreview (prevhost.exe)
                bool isPreviewHost = System.Diagnostics.Process.GetCurrentProcess().ProcessName.ToLowerInvariant().Contains("prevhost");
                
                if (pstream is System.Runtime.InteropServices.ComTypes.IStream com)
                {
                    _comStream = pstream;
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    using (var wrapper = new BufferedStream(new ComStreamWrapper(com, leaveOpen: true), 32 * 1024))
                    {
                        var (rows, info) = FitsPreviewControl.ParseFitsStream(wrapper);
                        _keywords = rows;
                        _metadata = info;
                        sw.Stop();
                        
                        // DETECTION: If parsing 3KB of header took > 200ms, we are on a slow network/cloud link.
                        // We will skip heavy thumbnail generation to keep Explorer responsive.
                        _isSlowLink = sw.ElapsedMilliseconds > 200;
                        
                        Log($"  Header parsed OK: {info.Width}x{info.Height} (Lat={sw.ElapsedMilliseconds}ms, Slow={_isSlowLink})");
                    }

                    // CRITICAL: We used to release the stream here for non-preview host usages (properties/thumbnails).
                    // BUT this broke thumbnails because GetThumbnail() needs that stream!
                    // Instead, we now rely on RenderThumbnail() or the destructor to release it.
                }
            }
            catch (Exception ex) { Log("IInitializeWithStream ERROR: " + ex); }
            finally { Log("IInitializeWithStream.END"); }
        }

        // ── IInitializeWithItem ─────────────────────────────────────────
        void IInitializeWithItem.Initialize(object psiObj, uint grfMode)
        {
            Log($"IInitializeWithItem — grfMode={grfMode}");
            try
            {
                if (psiObj is IShellItem psi)
                {
                    psi.GetDisplayName(0x80058000, out IntPtr ppszName);
                    if (ppszName != IntPtr.Zero)
                    {
                        _filePath = Marshal.PtrToStringAuto(ppszName);
                        Marshal.FreeCoTaskMem(ppszName);
                        Log($"  Path: '{_filePath}'");
                        Initialize(_filePath, grfMode);
                    }
                }
            }
            catch (Exception ex) { Log("IInitializeWithItem EXCEPTION: " + ex); }
        }

        // ── GetMetadata ─────────────────────────────────────────────────
        private ImageInfo? GetMetadata()
        {
            if (_metadata.HasValue) return _metadata;
            if (!string.IsNullOrEmpty(_filePath))
            {
                _metadata = ParseHeaderFromFile(_filePath);
            }
            return _metadata;
        }

        // ── IPreviewHandler ─────────────────────────────────────────────
        public void SetWindow(IntPtr hwnd, ref RECT rect)
        {
            Log($"SetWindow — 0x{hwnd:X}");
            _parentHwnd = hwnd;
            _bounds     = rect;
        }

        public void SetRect(ref RECT rect)
        {
            _bounds = rect;
            var ctrl = _control;
            if (ctrl != null && ctrl.IsHandleCreated && !ctrl.IsDisposed)
            {
                var r = new Rectangle(0, 0, rect.right - rect.left, rect.bottom - rect.top);
                ctrl.BeginInvoke(new Action(() =>
                {
                    try { if (!ctrl.IsDisposed) ctrl.Bounds = r; }
                    catch (Exception ex) { Log($"SetRect EXCEPTION: {ex.Message}"); }
                }));
            }
        }

        [DllImport("user32.dll")] static extern IntPtr SetParent(IntPtr c, IntPtr p);
        [DllImport("user32.dll")] static extern bool   SetWindowPos(IntPtr h, IntPtr i, int x, int y, int cx, int cy, uint f);
        [DllImport("user32.dll")] static extern bool   GetClientRect(IntPtr h, out RECT r);
        [DllImport("user32.dll", SetLastError = true)] static extern int GetWindowLong(IntPtr h, int n);
        [DllImport("user32.dll", SetLastError = true)] static extern int SetWindowLong(IntPtr h, int n, int v);
        const int GWL_STYLE = -16, WS_CHILD = 0x40000000, WS_VISIBLE = 0x10000000;

        private void TryBeginInvoke(Control c, Action a) { try { if (!c.IsDisposed) c.BeginInvoke(a); } catch { } }

        public void DoPreview()
        {
            Log($"DoPreview.BEGIN — path='{_filePath}' hwnd=0x{_parentHwnd:X}");
            try
            {
                if (_uiThread != null && _uiThread.IsAlive) { Log("  Already running, aborting NEW thread"); return; }

                if (_bounds.right - _bounds.left <= 0 || _bounds.bottom - _bounds.top <= 0)
                {
                    Log("  Bounds empty, querying parent...");
                    if (GetClientRect(_parentHwnd, out RECT pr)) _bounds = pr;
                }

                Rectangle bounds = new Rectangle(_bounds.left, _bounds.top,
                    _bounds.right - _bounds.left, _bounds.bottom - _bounds.top);
                Log($"  Final target bounds: {bounds}");

                Stream renderStream = null;
                if (!string.IsNullOrEmpty(_filePath))
                {
                    try
                    {
                        renderStream = new FileStream(_filePath, FileMode.Open, FileAccess.Read,
                                                      FileShare.ReadWrite | FileShare.Delete);
                        Log("  FileStream opened for render");
                    }
                    catch (Exception ex) { Log($"  FileStream failed: {ex.Message}"); }
                }

                string   fileName   = !string.IsNullOrEmpty(_filePath)
                    ? System.IO.Path.GetFileName(_filePath) : "FITS Stream";
                IntPtr   parentHost  = _parentHwnd;
                var      ready       = new System.Threading.ManualResetEventSlim(false);

                _uiThread = new System.Threading.Thread(() =>
                {
                    Log("UI Thread — started");
                    try
                    {
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);

                        var ctrl = new FitsPreviewControl();
                        _control = ctrl;
                        ctrl.Bounds = bounds;

                        var hwnd = ctrl.Handle;
                        Log($"UI Thread — HWND 0x{hwnd:X}");
                        ready.Set();

                        SetParent(hwnd, parentHost);
                        int st = GetWindowLong(hwnd, GWL_STYLE);
                        SetWindowLong(hwnd, GWL_STYLE, (st | WS_CHILD | WS_VISIBLE) & ~0x00C00000);
                        SetWindowPos(hwnd, IntPtr.Zero, bounds.X, bounds.Y,
                                     bounds.Width, bounds.Height, 0x0020 | 0x0040);
                        ctrl.Show();
                        ctrl.Update();

                        if (renderStream != null)
                        {
                            ctrl.LoadFits(renderStream, fileName, ownsStream: true, preParsedKeywords: _keywords, preParsedInfo: _metadata);
                        }
                        else if (_comStream is System.Runtime.InteropServices.ComTypes.IStream comStream)
                        {
                            var s = new BufferedStream(new ComStreamWrapper(comStream, leaveOpen: true), 32 * 1024);
                            ctrl.LoadFits(s, fileName, ownsStream: true, preParsedKeywords: _keywords, preParsedInfo: _metadata);
                        }
                        else if (!ctrl.IsDisposed)
                        {
                            Log("UI Thread — WARNING: no render stream");
                        }

                        Log("UI Thread — entering Application.Run");
                        Application.Run(new ApplicationContext());
                        Log("UI Thread — Application.Run returned");
                    }
                    catch (Exception ex)
                    {
                        Log("UI Thread EXCEPTION: " + ex);
                        try { renderStream?.Dispose(); } catch { }
                        ready.Set();
                    }
                });
                _uiThread.IsBackground = true;
                _uiThread.SetApartmentState(System.Threading.ApartmentState.STA);
                _uiThread.Start();
                bool ok = ready.Wait(3000);
                Log($"DoPreview.END — signalled={ok}");
            }
            catch (Exception ex) { Log("DoPreview EXCEPTION: " + ex); }
        }

        public void Unload()
        {
            Log("Unload.BEGIN");
            try
            {
                if (_control != null)
                {
                    var ctrl = _control;
                    _control = null;

                    TryBeginInvoke(ctrl, () =>
                    {
                        try { ctrl.Dispose(); }
                        catch (Exception ex) { Log("  Unload dispose EXCEPTION: " + ex); }
                        finally { try { Application.ExitThread(); } catch { } }
                    });
                }
                
                ReleaseStream();
            }
            catch (Exception ex) { Log("Unload EXCEPTION: " + ex); }
            finally { Log("Unload.END"); }
        }

        public void SetFocus()
        {
            if (_control != null && _control.IsHandleCreated)
                TryBeginInvoke(_control, () => _control.Focus());
        }

        public void QueryFocus(out IntPtr phwnd) =>
            phwnd = _control?.IsHandleCreated == true ? _control.Handle : IntPtr.Zero;

        public uint TranslateAccelerator(ref MSG pmsg) => 1;

        private object _site;
        public void SetSite(object pUnkSite) { _site = pUnkSite; }
        public void GetSite(ref Guid riid, out object ppvSite) { ppvSite = _site; }

        // ── IThumbnailProvider ──────────────────────────────────────────
        public void GetThumbnail(uint cx, out IntPtr hbmp, out WTS_ALPHATYPE pdwAlpha)
        {
            hbmp     = IntPtr.Zero;
            pdwAlpha = WTS_ALPHATYPE.WTSAT_RGB;
            Log($"GetThumbnail.BEGIN — cx={cx}");
            
            Bitmap bmp = null;
            try
            {
                var metaP = GetMetadata();
                if (!metaP.HasValue) { Log("  No metadata available for thumbnail"); return; }
                
                var info = metaP.Value;
                int size = (int)Math.Max(cx, 32);

                // PERFORMANCE: If we detected a slow link during Initialize, skip the image and show badge immediately.
                if (Settings.ShowImage && info.HasImage && !_isSlowLink)
                {
                    Stream stream = null;
                    bool owns = false;
                    try
                    {
                        if (!string.IsNullOrEmpty(_filePath))
                        {
                            stream = new FileStream(_filePath, FileMode.Open, FileAccess.Read, 
                                                   FileShare.ReadWrite | FileShare.Delete);
                            owns = true;
                        }
                        else if (_comStream is System.Runtime.InteropServices.ComTypes.IStream com)
                        {
                            stream = new BufferedStream(new ComStreamWrapper(com, leaveOpen: true), 32 * 1024);
                            owns = true;
                        }

                        if (stream != null)
                        {
                            Log($"  Rendering thumbnail {size}x{size} (1.5s timeout)...");
                            
                            // Watchdog: If network is too slow, bail to avoid locking Explorer
                            var renderTask = Task.Run(() => FitsPreviewControl.RenderThumbnail(stream, info, size, ownsStream: true));
                            
                            if (renderTask.Wait(1500)) 
                            {
                                bmp = renderTask.Result;
                                _comStream = null; // Stream already released by the render call
                            }
                            else 
                            {
                                Log("  Thumbnail TIMEOUT - Aborting I/O to release file handle");
                                // Forcing disposal here to break any pending network I/O
                                if (owns && stream != null) try { stream.Dispose(); } catch {}
                                ReleaseStream();
                            }
                            Log($"  Thumbnail result: {(bmp == null ? "FAILED/TIMEOUT" : "OK")}");
                        }
                    }
                    catch (Exception ex) { Log($"  Thumbnail stream render ERROR: {ex.Message}"); }
                    finally { 
                        if (owns && stream != null) try { stream.Dispose(); } catch {}
                        if (_comStream != null) ReleaseStream(); 
                    }
                }

                if (bmp == null)
                {
                    Log("  Using static badge for thumbnail");
                    bmp = FitsPreviewControl.RenderStaticBadge(info, size);
                }

                if (bmp != null)
                {
                    hbmp = bmp.GetHbitmap(Color.Black);
                }
            }
            catch (Exception ex) { Log("GetThumbnail EXCEPTION: " + ex); }
            finally 
            { 
                if (bmp != null) try { bmp.Dispose(); } catch {} 
                Log("GetThumbnail.END");
            }
        }

        // ── IPropertyStore ──────────────────────────────────────────────
        public uint GetCount(out uint cProps) { cProps = 7; return 0; }

        public uint GetAt(uint iProp, out PropertyKey pkey)
        {
            if      (iProp == 0) pkey = PKEYs.Subject;
            else if (iProp == 1) pkey = PKEYs.Photo_CameraModel;
            else if (iProp == 2) pkey = PKEYs.Photo_Exposure;
            else if (iProp == 3) pkey = PKEYs.Image_Width;
            else if (iProp == 4) pkey = PKEYs.Image_Height;
            else if (iProp == 5) pkey = PKEYs.Image_BitDepth;
            else if (iProp == 6) pkey = PKEYs.Category;
            else { pkey = new PropertyKey(); return 1; }
            return 0;
        }

        public uint GetValue(ref PropertyKey key, out PropVariant pv)
        {
            pv = new PropVariant();
            var metaP = GetMetadata();
            if (!metaP.HasValue) return 0x80004005;

            var meta = metaP.Value;
            if      (key.Equals(PKEYs.Subject))           pv.SetString(meta.Object ?? "");
            else if (key.Equals(PKEYs.Photo_CameraModel)) pv.SetString(meta.Instrument ?? meta.Camera ?? "");
            else if (key.Equals(PKEYs.Image_Width))       pv.SetInt(meta.Width);
            else if (key.Equals(PKEYs.Image_Height))      pv.SetInt(meta.Height);
            else if (key.Equals(PKEYs.Image_BitDepth))    pv.SetInt(Math.Abs(meta.BitPix));
            else if (key.Equals(PKEYs.Photo_Exposure))    pv.SetDouble(meta.Exposure);
            else if (key.Equals(PKEYs.Category))
            {
                string t = (meta.ImageType ?? "").ToLowerInvariant();
                pv.SetString(t.Contains("light") ? "Light" : t.Contains("flat") ? "Flat" : t.Contains("dark") ? "Dark" : t.Contains("bias") ? "Bias" : meta.ImageType ?? "");
            }
            return 0;
        }

        public uint SetValue(ref PropertyKey key, ref PropVariant pv) => 0x80030001;
        public uint Commit() => 0;

        #region Registration

        [ComRegisterFunction]
        public static void Register(Type t)
        {
            try
            {
                string guid  = t.GUID.ToString("B").ToUpper();
                string appid = "{6d2b5079-2f0b-48dd-ab7f-97cec514d30b}";
                LogStatic($"--- Register {guid} ---");

                using (RegistryKey key = Registry.LocalMachine.CreateSubKey(Settings.REG_PATH))
                    if (key != null)
                    {
                        if (key.GetValue(Settings.VAL_SHOW_IMAGE)   == null) key.SetValue(Settings.VAL_SHOW_IMAGE,    1,  RegistryValueKind.DWord);
                        if (key.GetValue(Settings.VAL_ENABLE_LOG)   == null) key.SetValue(Settings.VAL_ENABLE_LOG,    0,  RegistryValueKind.DWord);
                    }

                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".fits"))
                    key.SetValue("PerceivedType", "image");

                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\SystemFileAssociations\\.fits"))
                {
                    key.SetValue("FullDetails",    "prop:System.PropGroup.Image;System.Image.HorizontalSize;System.Image.VerticalSize;System.Image.BitDepth;System.PropGroup.Camera;System.Photo.CameraModel;System.Photo.ExposureTime;System.PropGroup.Description;System.Subject;System.Category");
                    key.SetValue("InfoTip",        "prop:System.ItemType;System.Size;System.Subject;System.Photo.CameraModel;System.Photo.ExposureTime;System.Category");
                    key.SetValue("PreviewDetails", "prop:*System.Image.HorizontalSize;*System.Image.VerticalSize;*System.Photo.CameraModel;*System.Photo.ExposureTime;*System.Subject;*System.Category");
                }

                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".fits\\ShellEx\\{8895b1c6-b41f-4c1c-a562-0d564250836f}"))
                    key.SetValue("", guid);
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".fits\\ShellEx\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}"))
                    key.SetValue("", guid);
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\SystemFileAssociations\\.fits\\ShellEx\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}"))
                    key.SetValue("", guid);
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".fits\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}"))
                    key.SetValue("", guid);
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\SystemFileAssociations\\.fits\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}"))
                    key.SetValue("", guid);

                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PreviewHandlers"))
                    key.SetValue(guid, "Fits Preview Handler");
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\.fits"))
                    key.SetValue("", guid);

                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\" + guid))
                {
                    key.SetValue("AppID", appid);
                }
                LogStatic("  Register OK");
            }
            catch (Exception ex) { LogStatic("  Register ERROR: " + ex); throw; }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            try
            {
                string guid = t.GUID.ToString("B").ToUpper();
                LogStatic($"--- Unregister {guid} ---");
                Registry.ClassesRoot.DeleteSubKeyTree(".fits\\ShellEx\\{8895b1c6-b41f-4c1c-a562-0d564250836f}", false);
                Registry.ClassesRoot.DeleteSubKeyTree(".fits\\ShellEx\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}", false);
                Registry.ClassesRoot.DeleteSubKeyTree(".fits\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}", false);
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PreviewHandlers", true))
                    key?.DeleteValue(guid, false);
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\.fits", false);
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey("CLSID\\" + guid, true))
                    key?.DeleteValue("AppID", false);
                LogStatic("  Unregister OK");
            }
            catch (Exception ex) { LogStatic("  Unregister ERROR: " + ex); }
        }

        #endregion
    }
}
