using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Generic;
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
        // It is never re-read after Initialize returns. All IPropertyStore and
        // IThumbnailProvider (badge) calls read from this cache only — zero I/O.
        private ImageInfo? _metadata;
        private List<(string, string, string)> _keywords;

        // Render stream used exclusively by DoPreview's render pipeline.
        // Lifetime: created in DoPreview → passed to LoadFits(ownsStream=true)
        //           → disposed by the render task's finally block as soon as
        //           pixel sampling completes (success, cancel, or error).
        // After that point no file handle is held, so the Shell can open the
        // file freely for the Details pane, Alt+Enter Properties, etc.
        //
        // COM IStream stored if InitializeWithStream is used.
        // It is held until DoPreview consumes it or Unload disposes it.
        private object _comStream;

        // ── Logging ─────────────────────────────────────────────────────
        private static void Log(string msg, bool force = false)
        {
            if (!force && !Settings.EnableTracing) return;
            msg = $"[{DateTime.Now:HH:mm:ss.fff}] [Ext] [T{System.Threading.Thread.CurrentThread.ManagedThreadId}] {msg}";
            System.Diagnostics.Debug.WriteLine(msg);
            try { File.AppendAllText(FitsPreviewControl.LogPath, msg + "\n"); } catch { }
        }

        static FitsPreviewHandlerExtension()
        {
            Log("--- CLASS LOADED ---", true);
            AppDomain.CurrentDomain.UnhandledException +=
                (s, e) => Log("!!! UNHANDLED: " + e.ExceptionObject, true);
        }

        public FitsPreviewHandlerExtension() => Log("=== constructor ===");

        // ── ParseHeaderFromFile ─────────────────────────────────────────
        // Opens the file, reads only until the END keyword (a few KB), closes.
        // File handle lifetime < 1 ms. Folder rename is never blocked.
        private static ImageInfo? ParseHeaderFromFile(string path)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read,
                                               FileShare.ReadWrite | FileShare.Delete))
                {
                    var (_, info) = FitsPreviewControl.ParseFitsStream(fs);
                    Log($"ParseHeaderFromFile — {info.Width}x{info.Height} BITPIX={info.BitPix}");
                    return info;
                }
            }
            catch (Exception ex) { Log($"ParseHeaderFromFile failed: {ex.Message}"); return null; }
        }

        // ── IInitializeWithFile ─────────────────────────────────────────
        // Called by: Explorer preview pane (file path available), SearchIndexer,
        // Details pane, InfoTip. Most common activation path.
        public void Initialize(string pszFilePath, uint grfMode)
        {
            Log($"IInitializeWithFile — '{pszFilePath}'");
            _filePath    = pszFilePath;
            _comStream   = null;
            _keywords    = null;
            _metadata    = null;

            try {
                using (var fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
                    var (rows, info) = FitsPreviewControl.ParseFitsStream(fs);
                    _keywords = rows;
                    _metadata = info;
                }
            } catch (Exception ex) { Log($"IInitializeWithFile Header Read failed: {ex.Message}"); }
        }

        // ── IInitializeWithStream ───────────────────────────────────────
        // Called by: Explorer preview pane on virtual / network filesystems
        // where only a stream (no path) is available.
        //
        // We parse the header and — if image rendering is enabled — pre-buffer
        // only the stride-sampled pixel rows into a MemoryStream. Then we
        // release the COM IStream immediately, before this method returns.
        // The Shell is never blocked waiting for us to release the file.
        void IInitializeWithStream.Initialize(object pstream, uint grfMode)
        {
            Log($"IInitializeWithStream — grfMode={grfMode}");
            _metadata  = null;
            _comStream = null; 

            if (pstream is System.Runtime.InteropServices.ComTypes.IStream comStream)
            {
                try
                {
                    // 1. Fast header-only parse (fast, no lock held yet).
                    using (var wrapper = new ComStreamWrapper(comStream, leaveOpen: true))
                    {
                        wrapper.Seek(0, SeekOrigin.Begin);
                        var (rows, info) = FitsPreviewControl.ParseFitsStream(wrapper);
                        _keywords = rows;
                        _metadata = info;
                        
                        // 2. Buffer only the header and stride-sampled pixels into a SampledStream.
                        //    Then we release the COM stream immediately.
                        int targetDim = Settings.ShowImage ? 600 : -1;
                        var ss = new SampledStream(info.Width, info.Height, info.BitPix, info.DataOffset);
                        
                        // Buffer header
                        wrapper.Seek(0, SeekOrigin.Begin);
                        byte[] hbuf = new byte[info.DataOffset > 0 ? info.DataOffset : 2880];
                        wrapper.Read(hbuf, 0, hbuf.Length);
                        ss.WriteSample(0, hbuf);

                        if (targetDim > 0 && info.HasImage)
                        {
                            // Stride logic duplicated from RenderImage for pre-buffering
                            double aspect = (double)info.Width / info.Height;
                            int outW = info.Width >= info.Height ? Math.Min(info.Width, targetDim) : Math.Max(1, (int)(Math.Min(info.Height, targetDim) * aspect));
                            int outH = info.Width >= info.Height ? Math.Max(1, (int)(outW / aspect)) : Math.Min(info.Height, targetDim);
                            bool isBayer = !string.IsNullOrEmpty(info.BayerPattern) && info.Planes == 1;
                            int sy = Math.Max(1, info.Height / outH);
                            if (isBayer) sy = (sy / 2) * 2; // Force even stride to keep Bayer pairs together
                            if (sy < 1) sy = 1;

                            int bpp = info.BytesPerPixel;
                            long rowBytes = (long)info.Width * bpp;

                            byte[] row = new byte[rowBytes];
                            for (int p = 0; p < info.Planes; p++)
                            {
                                long planeOff = info.DataOffset + (long)p * info.Height * rowBytes;
                                for (int y = 0; y <= info.Height - (isBayer ? 2 : 1); y += sy)
                                {
                                    wrapper.Seek(planeOff + (long)y * rowBytes, SeekOrigin.Begin);
                                    wrapper.Read(row, 0, (int)rowBytes);
                                    ss.WriteSample(planeOff + (long)y * rowBytes, row);
                                    if (isBayer) {
                                        wrapper.Read(row, 0, (int)rowBytes);
                                        ss.WriteSample(planeOff + (long)(y+1) * rowBytes, row);
                                    }
                                }
                            }
                        }
                        ss.Position = 0;
                        _comStream = ss; // Store our in-memory SampledStream instead of COM object
                        Log($"  SampledStream buffered: {ss.BufferedBytes / 1024} KB. Header + {ss.SampleCount} rows.");
                    }
                    // Release the COM IStream now so Shell can unlock the file
                    Marshal.ReleaseComObject(pstream);
                    Log("  COM IStream released.");
                }
                catch (Exception ex) { Log($"IInitializeWithStream EXCEPTION: {ex}"); }
            }
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

        // Not used anymore as we now stream directly from IStream/FileStream
        // to avoid memory pressure and delay.

        // ── GetMetadata ─────────────────────────────────────────────────
        // Pure cache read. No I/O. Safe from any thread at any time.
        private ImageInfo? GetMetadata()
        {
            if (_metadata.HasValue) return _metadata;
            // Emergency fallback (should not be needed in normal operation).
            if (!string.IsNullOrEmpty(_filePath))
            {
                Log("GetMetadata — cache miss, emergency fallback");
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

        public void DoPreview()
        {
            Log($"DoPreview — path='{_filePath}' hwnd=0x{_parentHwnd:X}");
            try
            {
                if (_uiThread != null && _uiThread.IsAlive) { Log("DoPreview — already running"); return; }

                if (_bounds.right - _bounds.left <= 0 || _bounds.bottom - _bounds.top <= 0)
                    if (GetClientRect(_parentHwnd, out RECT pr)) _bounds = pr;

                Rectangle bounds = new Rectangle(_bounds.left, _bounds.top,
                    _bounds.right - _bounds.left, _bounds.bottom - _bounds.top);

                Stream renderStream = null;
                if (_comStream is Stream s)
                {
                    renderStream = s;
                    _comStream = null;
                }
                else if (_comStream is System.Runtime.InteropServices.ComTypes.IStream comStream)
                {
                    renderStream = new ComStreamWrapper(comStream);
                    _comStream = null; // transfer ownership to wrapper
                }

                if (renderStream == null && !string.IsNullOrEmpty(_filePath))
                {
                    try
                    {
                        renderStream = new FileStream(_filePath, FileMode.Open, FileAccess.Read,
                                                      FileShare.ReadWrite | FileShare.Delete);
                        Log("DoPreview — FileStream opened for render");
                    }
                    catch (Exception ex) { Log($"DoPreview — FileStream failed: {ex.Message}"); }
                }

                string   fileName   = !string.IsNullOrEmpty(_filePath)
                    ? System.IO.Path.GetFileName(_filePath) : "FITS Stream";
                Stream   localStream  = renderStream;
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
                        ctrl.Update(); // Force initial paint of empty container

                        if (localStream != null && !ctrl.IsDisposed)
                        {
                            // We use BeginInvoke to ensure the Application.Run message loop is 
                            // already active. This allows the grid to be populated and painted
                            // BEFORE the background image task starts saturating the IO.
                            ctrl.BeginInvoke(new Action(() => {
                                ctrl.LoadFits(localStream, fileName, ownsStream: true, preParsedKeywords: _keywords, preParsedInfo: _metadata);
                            }));
                        }
                        else if (!ctrl.IsDisposed)
                        {
                            Log("UI Thread — no render stream (metadata-only or ShowImage=false)");
                        }

                        Application.Run(new ApplicationContext());
                        Log("UI Thread — pump exited");
                    }
                    catch (Exception ex)
                    {
                        Log("UI Thread EXCEPTION: " + ex);
                        // Dispose stream if LoadFits never got to take ownership.
                        try { localStream?.Dispose(); } catch { }
                        ready.Set();
                    }
                });
                _uiThread.IsBackground = true;
                _uiThread.SetApartmentState(System.Threading.ApartmentState.STA);
                _uiThread.Start();
                Log($"DoPreview — signalled={ready.Wait(3000)}");
            }
            catch (Exception ex) { Log("DoPreview EXCEPTION: " + ex); }
        }

        public void Unload()
        {
            Log("Unload — called");
            var th = _uiThread;
            _uiThread = null;

            if (_control != null && _control.IsHandleCreated)
            {
                var ctrl = _control;
                _control = null;
                var done = new System.Threading.ManualResetEventSlim(false);
                try
                {
                    ctrl.BeginInvoke(new Action(() =>
                    {
                        try   { ctrl.Dispose(); }
                        catch (Exception ex) { Log("Unload dispose EXCEPTION: " + ex); }
                        finally { try { Application.ExitThread(); } catch { } done.Set(); }
                    }));
                }
                catch { done.Set(); }
                done.Wait(3000);
            }

            if (th != null && th.IsAlive) th.Join(2000);

            if (_comStream != null)
            {
                try { Marshal.ReleaseComObject(_comStream); } catch { }
                _comStream = null;
                Log("Unload — COM stream released");
            }

            Log("Unload — done");
        }

        public void SetFocus()
        {
            if (_control != null && _control.IsHandleCreated)
                _control.BeginInvoke(new Action(() => _control.Focus()));
        }

        public void QueryFocus(out IntPtr phwnd) =>
            phwnd = _control?.IsHandleCreated == true ? _control.Handle : IntPtr.Zero;

        public uint TranslateAccelerator(ref MSG pmsg) => 1;

        private object _site;
        public void SetSite(object pUnkSite) { _site = pUnkSite; }
        public void GetSite(ref Guid riid, out object ppvSite) { ppvSite = _site; }

        // ── IThumbnailProvider ──────────────────────────────────────────
        // Always uses its own short-lived stream. Never touches _renderStream.
        public void GetThumbnail(uint cx, out IntPtr hbmp, out WTS_ALPHATYPE pdwAlpha)
        {
            hbmp     = IntPtr.Zero;
            pdwAlpha = WTS_ALPHATYPE.WTSAT_RGB;
            Log($"GetThumbnail — cx={cx}");
            
            System.Drawing.Bitmap bmp = null;
            try
            {
                var metaP = GetMetadata();
                if (!metaP.HasValue) { Log("GetThumbnail — no metadata"); return; }
                var info = metaP.Value;
                int size = (int)Math.Max(cx, 32);

                bool tryImage = Settings.ShowImage && info.HasImage;
                
                if (tryImage)
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
                        else if (_comStream is SampledStream s)
                        {
                            stream = s;
                            // SampledStream implements Seek/Position, and we created it, 
                            // so we can read from it directly.
                        }
                        else if (_comStream is System.Runtime.InteropServices.ComTypes.IStream com)
                        {
                            // Fallback for raw COM streams
                            stream = new ComStreamWrapper(com, leaveOpen: true);
                            owns = true;
                        }

                        if (stream != null)
                        {
                            bmp = FitsPreviewControl.RenderThumbnail(stream, info, size);
                            if (bmp != null) Log("GetThumbnail — Image rendered successfully");
                        }
                    }
                    catch (Exception ex) { Log($"GetThumbnail Image Render failed: {ex.Message}"); }
                    finally { if (owns && stream != null) stream.Dispose(); }
                }

                // Fallback to badge if image failed or is disabled
                if (bmp == null)
                {
                    Log("GetThumbnail — Rendering static badge");
                    bmp = FitsPreviewControl.RenderStaticBadge(info, size);
                }

                if (bmp != null)
                {
                    hbmp = bmp.GetHbitmap(System.Drawing.Color.Black);
                    // Standard thumbnails are non-alpha for shell efficiency
                    pdwAlpha = WTS_ALPHATYPE.WTSAT_RGB;
                }
            }
            catch (Exception ex) { Log("GetThumbnail — FATAL EXCEPTION: " + ex); }
            finally { if (bmp != null) bmp.Dispose(); }
        }

        // ── IPropertyStore ──────────────────────────────────────────────
        // Zero I/O after Initialize. Reads only from _metadata.

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
            if (!metaP.HasValue) return 0x80004005; // E_FAIL
            var meta = metaP.Value;

            if      (key.Equals(PKEYs.Subject))           pv.SetString(meta.Object ?? "");
            else if (key.Equals(PKEYs.Photo_CameraModel)) pv.SetString(meta.Instrument ?? meta.Camera ?? "");
            else if (key.Equals(PKEYs.Image_Width))       pv.SetInt(meta.Width);
            else if (key.Equals(PKEYs.Image_Height))      pv.SetInt(meta.Height);
            else if (key.Equals(PKEYs.Image_BitDepth))    pv.SetInt(Math.Abs(meta.BitPix));
            else if (key.Equals(PKEYs.Photo_Exposure))    pv.SetDouble(meta.Exposure);
            else if (key.Equals(PKEYs.Category))
            {
                string t   = (meta.ImageType ?? "").ToLowerInvariant();
                string cat = t.Contains("light") ? "Light"
                           : t.Contains("flat")  ? "Flat"
                           : t.Contains("dark")  ? "Dark"
                           : t.Contains("bias")  ? "Bias"
                           : meta.ImageType ?? "";
                pv.SetString(cat);
            }
            return 0; // S_OK
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
                Log($"--- Register {guid} ---", true);

                using (RegistryKey key = Registry.LocalMachine.CreateSubKey(Settings.REG_PATH))
                    if (key != null)
                    {
                        if (key.GetValue(Settings.VAL_SHOW_IMAGE)   == null) key.SetValue(Settings.VAL_SHOW_IMAGE,    1,  RegistryValueKind.DWord);
                        if (key.GetValue(Settings.VAL_ENABLE_LOG)   == null) key.SetValue(Settings.VAL_ENABLE_LOG,    0,  RegistryValueKind.DWord);
                        if (key.GetValue(Settings.VAL_SPLITTER_POS) == null) key.SetValue(Settings.VAL_SPLITTER_POS, -1,  RegistryValueKind.DWord);
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
                    key.DeleteValue("ManualSafeSave",          false);
                    key.DeleteValue("DisableProcessIsolation", false);
                }
                Log("  Register OK", true);
            }
            catch (Exception ex) { Log("  Register ERROR: " + ex, true); throw; }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            try
            {
                string guid = t.GUID.ToString("B").ToUpper();
                Log($"--- Unregister {guid} ---", true);
                Registry.ClassesRoot.DeleteSubKeyTree(".fits\\ShellEx\\{8895b1c6-b41f-4c1c-a562-0d564250836f}", false);
                Registry.ClassesRoot.DeleteSubKeyTree(".fits\\ShellEx\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}", false);
                Registry.ClassesRoot.DeleteSubKeyTree(".fits\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}", false);
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PreviewHandlers", true))
                    key?.DeleteValue(guid, false);
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\.fits", false);
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey("CLSID\\" + guid, true))
                    key?.DeleteValue("AppID", false);
                Log("  Unregister OK", true);
            }
            catch (Exception ex) { Log("  Unregister ERROR: " + ex, true); }
        }

        #endregion
    }
}
