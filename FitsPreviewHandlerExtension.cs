using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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

        public void SetString(string val)
        {
            vt = 31; // VT_LPWSTR
            ptr = Marshal.StringToCoTaskMemUni(val);
        }

        public void SetInt(int val)
        {
            vt = 3; // VT_I4
            iVal = val;
        }

        public void SetDouble(double val)
        {
            vt = 5; // VT_R8
            dblVal = val;
        }

        [DllImport("ole32.dll")]
        public static extern int PropVariantClear(ref PropVariant pvar);

        public void Dispose()
        {
            PropVariantClear(ref this);
        }
    }

    public static class PKEYs
    {
        public static readonly Guid PSG_SUMMARY = new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
        public static readonly Guid PSG_IMAGE   = new Guid("6444048F-4C8B-11D1-8B70-080036B11A03");
        public static readonly Guid PKEY_Photo  = new Guid("14B81DA1-0135-4D31-96D9-6CBFC9671A99");

        public static PropertyKey Subject           = new PropertyKey { fmtid = PSG_SUMMARY, pid = 3 };
        public static PropertyKey Image_Width       = new PropertyKey { fmtid = PSG_IMAGE, pid = 3 };
        public static PropertyKey Image_Height      = new PropertyKey { fmtid = PSG_IMAGE, pid = 4 };
        public static PropertyKey Image_BitDepth    = new PropertyKey { fmtid = PSG_IMAGE, pid = 7 };
        public static PropertyKey Photo_CameraModel = new PropertyKey { fmtid = PKEY_Photo, pid = 272 };
        public static PropertyKey Photo_Exposure    = new PropertyKey { fmtid = PKEY_Photo, pid = 33434 };
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("fc4801a3-2ba9-11cf-a229-00aa003d7352")]
    public interface IObjectWithSite
    {
        void SetSite([In, MarshalAs(UnmanagedType.IUnknown)] object pUnkSite);
        void GetSite(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvSite);
    }

    /// <summary>Alpha channel type reported back by <see cref="IThumbnailProvider.GetThumbnail"/>.</summary>
    public enum WTS_ALPHATYPE : int { WTSAT_UNKNOWN = 0, WTSAT_RGB = 1, WTSAT_ARGB = 2 }

    /// <summary>
    /// Shell thumbnail provider interface (IID {e357fccd-a995-4576-b01f-234630154e96}).
    /// Windows calls GetThumbnail on a worker thread; no STA/WinForms objects may be created here.
    /// </summary>
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("e357fccd-a995-4576-b01f-234630154e96")]
    public interface IThumbnailProvider
    {
        void GetThumbnail(uint cx, out IntPtr hbmp, out WTS_ALPHATYPE pdwAlpha);
    }

    [ComVisible(true)]
    [Guid("AF1C3D6A-81E9-4F5B-9A8C-2D9E71F04B3E")]
    [ClassInterface(ClassInterfaceType.None)]
    public class FitsPreviewHandlerExtension : IPreviewHandler, IInitializeWithFile, IInitializeWithStream, IInitializeWithItem, IObjectWithSite, IPropertyStore, IThumbnailProvider
    {
        private string _filePath;
        private IntPtr _parentHwnd;
        private RECT _bounds;
        private FitsPreviewControl _control;
        private System.Threading.Thread _uiThread;
        private Stream _stream; // Direct zero-copy stream
        private ImageInfo? _metadata; // Cached metadata for property store
        
        // Shared log writer — delegates to the same path used by FitsPreviewControl
        private static void Log(string msg, bool force = false)
        {
            if (!force && !Settings.EnableTracing) return;
            msg = $"[{DateTime.Now:HH:mm:ss.fff}] [Extension] [T{System.Threading.Thread.CurrentThread.ManagedThreadId}] {msg}";
            System.Diagnostics.Debug.WriteLine(msg);
            try
            {
                File.AppendAllText(FitsPreviewControl.LogPath, msg + "\n");
            }
            catch { /* fallback only */ }
        }

        static FitsPreviewHandlerExtension()
        {
            // The earliest possible trace point in the .NET runtime
            Log("--- STATIC CONSTRUCTOR INVOKED (Class loaded into AppDomain) ---");
            Log($"  Domain: {AppDomain.CurrentDomain.FriendlyName}");
            
            AppDomain.CurrentDomain.UnhandledException += (s, e) => {
                Log("!!! UNHANDLED EXCEPTION in AppDomain: " + e.ExceptionObject);
            };
        }

        public FitsPreviewHandlerExtension()
        {
            Log("=== FitsPreviewHandlerExtension constructor invoked ===");
            Log($"  Assembly: {System.Reflection.Assembly.GetExecutingAssembly().Location}");
            Log($"  LogPath : {FitsPreviewControl.LogPath}");
        }

        public void Initialize(string pszFilePath, uint grfMode)
        {
            Log($"IInitializeWithFile.Initialize — path='{pszFilePath}' grfMode={grfMode}");
            _filePath = pszFilePath;
            _metadata = null;
        }

        void IInitializeWithStream.Initialize(object pstream, uint grfMode)
        {
            Log($"IInitializeWithStream.Initialize — grfMode={grfMode}, stream is {pstream?.GetType().FullName ?? "null"}");
            try
            {
                if (pstream is System.Runtime.InteropServices.ComTypes.IStream comStream)
                {
                    Log($"IInitializeWithStream.Initialize — ready for Zero-Copy");
                    _stream = new ComStreamWrapper(comStream);
                }
                else
                {
                    Log("  WARNING: pstream does not implement IStream");
                }
                _metadata = null;
            }
            catch (Exception ex) { Log("IInitializeWithStream EXCEPTION: " + ex); }
        }

        void IInitializeWithItem.Initialize(object psiObj, uint grfMode)
        {
            Log($"IInitializeWithItem.Initialize — grfMode={grfMode}, item is {psiObj?.GetType().FullName ?? "null"}");
            try
            {
                if (psiObj is IShellItem psi)
                {
                    psi.GetDisplayName(0x80058000, out IntPtr ppszName);
                    if (ppszName != IntPtr.Zero)
                    {
                        _filePath = Marshal.PtrToStringAuto(ppszName);
                        Marshal.FreeCoTaskMem(ppszName);
                        Log($"  Got path from IShellItem: '{_filePath}'");
                    }
                    else
                    {
                        Log("  WARNING: GetDisplayName returned null pointer");
                    }
                }
                else
                {
                    Log("  WARNING: psiObj does not implement IShellItem");
                }
            }
            catch (Exception ex) { Log("  EXCEPTION in IInitializeWithItem.Initialize: " + ex); }
        }

        public void SetWindow(IntPtr hwnd, ref RECT rect)
        {
            Log($"SetWindow — hwnd=0x{hwnd:X} rect=[{rect.left},{rect.top},{rect.right},{rect.bottom}]");
            _parentHwnd = hwnd;
            _bounds = rect;
        }

        public void SetRect(ref RECT rect)
        {
            Log($"SetRect — rect=[{rect.left},{rect.top},{rect.right},{rect.bottom}]");
            _bounds = rect;
            if (_control != null && _control.IsHandleCreated)
            {
                var r = new Rectangle(0, 0, rect.right - rect.left, rect.bottom - rect.top);
                Log($"SetRect — control READY (HWND: {_control.Handle:X}), resizing to {r}");
                _control.BeginInvoke(new Action(() =>
                {
                    try { _control.Bounds = r; }
                    catch (Exception ex) { Log($"SetRect — EXCEPTION during resize: {ex.Message}"); }
                }));
            }
            else
            {
                Log($"SetRect — control not ready yet (_control={(_control == null ? "null" : "non-null")}, HandleCreated={_control?.IsHandleCreated ?? false})");
            }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        // WS_CHILD | WS_VISIBLE
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        private const int GWL_STYLE = -16;
        private const int WS_CHILD = 0x40000000;
        private const int WS_VISIBLE = 0x10000000;

        public void DoPreview()
        {
            Log($"DoPreview — _filePath='{_filePath}' parentHwnd=0x{_parentHwnd:X}");
            Log($"DoPreview — bounds=[{_bounds.left},{_bounds.top},{_bounds.right},{_bounds.bottom}]");
            try
            {
                if (_uiThread != null)
                {
                    Log("DoPreview — UI thread already running, skipping");
                    return;
                }

                string pathToLoad = _filePath;
                IntPtr parentHost = _parentHwnd;

                // Robust bounds detection: if Explorer gave [0,0,0,0], ask parent directly
                if (_bounds.right - _bounds.left <= 0 || _bounds.bottom - _bounds.top <= 0)
                {
                    Log("DoPreview — _bounds is empty! Getting client rect from parent.");
                    if (GetClientRect(parentHost, out RECT parentRect))
                    {
                        _bounds = parentRect;
                        Log($"DoPreview — GetClientRect(parent) → [{_bounds.left},{_bounds.top},{_bounds.right},{_bounds.bottom}]");
                    }
                }

                Rectangle initialBounds = new Rectangle(
                    _bounds.left, _bounds.top,
                    _bounds.right - _bounds.left,
                    _bounds.bottom - _bounds.top);

                Log($"DoPreview — final initialBounds={initialBounds}");

                // Synchronization: wait until the control handle exists
                var ready = new System.Threading.ManualResetEventSlim(false);

                _uiThread = new System.Threading.Thread(() =>
                {
                    Log("UI Thread — started");
                    try
                    {
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);
                        Log("UI Thread — Application styles set");

                        _control = new FitsPreviewControl();
                        _control.Bounds = initialBounds;
                        Log($"UI Thread — FitsPreviewControl created, bounds={_control.Bounds}");

                        // Force handle creation before signalling the caller
                        var hwndControl = _control.Handle;
                        Log($"UI Thread — HWND created: 0x{hwndControl:X}");
                        ready.Set();

                        // Reparent + style as WS_CHILD so it sits inside the Explorer preview pane
                        var prevParent = SetParent(hwndControl, parentHost);
                        Log($"UI Thread — SetParent → prevParent=0x{prevParent:X}");

                        int style = GetWindowLong(hwndControl, GWL_STYLE);
                        Log($"UI Thread — original WS style: 0x{style:X}");
                        int newStyle = (style | WS_CHILD | WS_VISIBLE) & ~0x00C00000; // drop WS_CAPTION
                        SetWindowLong(hwndControl, GWL_STYLE, newStyle);
                        Log($"UI Thread — new WS style: 0x{newStyle:X}");

                        bool swpOk = SetWindowPos(hwndControl, IntPtr.Zero,
                            initialBounds.X, initialBounds.Y,
                            initialBounds.Width, initialBounds.Height,
                            0x0020 | 0x0040); // SWP_FRAMECHANGED | SWP_NOACTIVATE
                        Log($"UI Thread — SetWindowPos result: {swpOk}");

                        _control.Show();
                        Log("UI Thread — control shown");

                        if (_stream != null)
                        {
                            Log("UI Thread — loading from direct STREAM");
                            _control.LoadFits(_stream, "FITS Stream");
                        }
                        else if (!string.IsNullOrEmpty(pathToLoad))
                        {
                            Log($"UI Thread — loading from PATH: '{pathToLoad}'");
                            _control.LoadFits(pathToLoad);
                        }
                        else
                        {
                            Log("UI Thread — WARNING: no stream or path to load");
                        }

                        // Keep the STA message pump alive indefinitely
                        Log("UI Thread — entering Application.Run(ApplicationContext)");
                        var ctx = new ApplicationContext();
                        Application.Run(ctx);
                        Log("UI Thread — Application.Run returned (unexpected)");
                    }
                    catch (Exception tEx)
                    {
                        Log("UI Thread — UNHANDLED EXCEPTION: " + tEx);
                        ready.Set(); // unblock caller even on failure
                    }
                });
                _uiThread.IsBackground = true;
                _uiThread.SetApartmentState(System.Threading.ApartmentState.STA);
                _uiThread.Start();
                Log("DoPreview — UI thread started, waiting for handle (max 3 s)...");

                bool signalled = ready.Wait(3000);
                Log($"DoPreview — ready.Wait returned: signalled={signalled} (HWND: {_control?.Handle:X})");
                if (!signalled) Log("DoPreview — ERROR: UI thread did not signal handle creation in time.");
            }
            catch (Exception ex) { Log("DoPreview — EXCEPTION: " + ex); }
        }

        public void Unload()
        {
            Log("Unload — called");
            if (_control != null && _control.IsHandleCreated)
            {
                Log("Unload — disposing control and exiting UI thread message loop");
                var controlToDispose = _control;
                controlToDispose.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        controlToDispose.Dispose();
                        Application.ExitThread();
                        Log("Unload — Application.ExitThread called inside UI thread");
                    }
                    catch (Exception ex) { Log("Unload — EXCEPTION during dispose: " + ex); }
                }));
                _control = null;
            }
            else
            {
                Log("Unload — control was null or handle not created");
            }
            _uiThread = null;
            Log("Unload — cleaning up");
            if (_stream != null)
            {
                try { _stream.Dispose(); } catch { }
                _stream = null;
            }
            Log("Unload — done");
        }

        public void SetFocus()
        {
            Log("SetFocus");
            if (_control != null && _control.IsHandleCreated)
                _control.BeginInvoke(new Action(() => _control.Focus()));
        }

        public void QueryFocus(out IntPtr phwnd)
        {
            IntPtr result = _control?.IsHandleCreated == true ? _control.Handle : IntPtr.Zero;
            Log($"QueryFocus — returning 0x{result:X}");
            phwnd = result;
        }

        public uint TranslateAccelerator(ref MSG pmsg) { return 1; /* S_FALSE = not handled */ }

        private object _site;
        public void SetSite(object pUnkSite) { _site = pUnkSite; }
        public void GetSite(ref Guid riid, out object ppvSite) { ppvSite = _site; }

        // ── IThumbnailProvider Implementation ────────────────────────────
        public void GetThumbnail(uint cx, out IntPtr hbmp, out WTS_ALPHATYPE pdwAlpha)
        {
            hbmp     = IntPtr.Zero;
            pdwAlpha = WTS_ALPHATYPE.WTSAT_RGB;
            Log($"IThumbnailProvider.GetThumbnail — cx={cx}");

            Stream streamToUse   = null;
            bool   shouldDispose = false;
            System.Drawing.Bitmap bmp = null;
            try
            {
                // Resolve stream: prefer the COM IStream already handed to us,
                // fall back to opening the file path directly.
                if (_stream != null)
                    streamToUse = _stream;
                else if (!string.IsNullOrEmpty(_filePath))
                {
                    streamToUse   = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    shouldDispose = true;
                }

                if (streamToUse == null)
                {
                    Log("GetThumbnail — no stream or path available");
                    return;
                }

                // Step 1: parse the FITS header only (reads until END; typically a few KB).
                // ParseFitsStream seeks to 0 internally, so stream position doesn't matter.
                var (_, info) = FitsPreviewControl.ParseFitsStream(streamToUse);
                _metadata = info; // opportunistic cache for concurrent IPropertyStore queries

                int size = (int)Math.Max(cx, 32);

                // Step 2: render either a real (stride-sampled) thumbnail or a static badge.
                if (Settings.ShowImage && info.HasImage)
                {
                    // Stride: e.g. 4656-wide sensor at cx=256 → stride≈18 → reads ~1/324 of the data.
                    Log($"GetThumbnail — stride-sampled render to {size}\u00d7{size}");
                    bmp = FitsPreviewControl.RenderThumbnail(streamToUse, info, size);
                }
                else
                {
                    // Static badge: zero pixel data read — header metadata only.
                    Log($"GetThumbnail — static badge (ShowImage={Settings.ShowImage}, HasImage={info.HasImage})");
                    bmp = FitsPreviewControl.RenderStaticBadge(info, size);
                }

                if (bmp != null)
                {
                    // Shell owns this HBITMAP and will free it via DeleteObject when done.
                    hbmp = bmp.GetHbitmap(System.Drawing.Color.Black);
                    Log($"GetThumbnail — success hbmp=0x{hbmp:X}");
                }
            }
            catch (Exception ex)
            {
                Log("GetThumbnail — EXCEPTION: " + ex);
            }
            finally
            {
                bmp?.Dispose();
                if (shouldDispose) streamToUse?.Dispose();
            }
        }

        // ── IPropertyStore Implementation ───────────────────────────────
        private ImageInfo? GetMetadata()
        {
            if (_metadata.HasValue) return _metadata;
            
            Stream streamToParse = null;
            bool shouldDispose = false;

            if (_stream != null)
                streamToParse = _stream;
            else if (!string.IsNullOrEmpty(_filePath))
            {
                streamToParse = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                shouldDispose = true;
            }

            if (streamToParse != null)
            {
                try 
                {
                    var (_, info) = FitsPreviewControl.ParseFitsStream(streamToParse);
                    _metadata = info;
                }
                finally { if (shouldDispose) streamToParse.Dispose(); }
            }
            
            return _metadata;
        }

        public uint GetCount(out uint cProps) 
        { 
            cProps = 6;
            Log("IPropertyStore.GetCount — returning 6");
            return 0; // S_OK
        }

        public uint GetAt(uint iProp, out PropertyKey pkey)
        {
            Log($"IPropertyStore.GetAt — requested index {iProp}");
            if (iProp == 0) pkey = PKEYs.Subject;
            else if (iProp == 1) pkey = PKEYs.Photo_CameraModel;
            else if (iProp == 2) pkey = PKEYs.Photo_Exposure;
            else if (iProp == 3) pkey = PKEYs.Image_Width;
            else if (iProp == 4) pkey = PKEYs.Image_Height;
            else if (iProp == 5) pkey = PKEYs.Image_BitDepth;
            else { pkey = new PropertyKey(); return 1; /* S_FALSE */ }
            return 0;
        }

        public uint GetValue(ref PropertyKey key, out PropVariant pv)
        {
            pv = new PropVariant();
            Log($"IPropertyStore.GetValue — requested {key.fmtid} pid={key.pid}");
            var metaP = GetMetadata();
            if (!metaP.HasValue) {
                Log("  GetMetadata() failed.");
                return 0x80004005; // E_FAIL
            }
            var meta = metaP.Value;

            if (key.Equals(PKEYs.Subject)) { pv.SetString(meta.Object ?? ""); Log($"  Returned Subject: {meta.Object}"); }
            else if (key.Equals(PKEYs.Photo_CameraModel)) { pv.SetString(meta.Instrument ?? meta.Camera ?? ""); Log($"  Returned Camera: {meta.Instrument ?? meta.Camera}"); }
            else if (key.Equals(PKEYs.Image_Width)) { pv.SetInt(meta.Width); Log($"  Returned Width: {meta.Width}"); }
            else if (key.Equals(PKEYs.Image_Height)) { pv.SetInt(meta.Height); Log($"  Returned Height: {meta.Height}"); }
            else if (key.Equals(PKEYs.Image_BitDepth)) { pv.SetInt(Math.Abs(meta.BitPix)); Log($"  Returned BitDepth: {meta.BitPix}"); }
            else if (key.Equals(PKEYs.Photo_Exposure)) { pv.SetDouble(meta.Exposure); Log($"  Returned Exposure: {meta.Exposure}"); }
            else { Log("  Returned EMPTY"); }
            
            return 0;
        }

        public uint SetValue(ref PropertyKey key, ref PropVariant pv) => 0x80030001; // STG_E_ACCESSDENIED
        public uint Commit() => 0;

        #region Registration Logic

        [ComRegisterFunction]
        public static void Register(Type t)
        {
            try
            {
                string guid = t.GUID.ToString("B").ToUpper();
                // Standard Microsoft AppID for Preview Host (prevhost.exe)
                string appid = "{6d2b5079-2f0b-48dd-ab7f-97cec514d30b}";
                
                Log($"--- Registering Preview Handler {guid} ---", true);

                // 0. Create Configuration keys in HKLM (Safe defaults for all users)
                try {
                    using (RegistryKey key = Registry.LocalMachine.CreateSubKey(Settings.REG_PATH))
                    {
                        if (key != null) {
                            if (key.GetValue(Settings.VAL_SHOW_IMAGE) == null) 
                                key.SetValue(Settings.VAL_SHOW_IMAGE, 1, RegistryValueKind.DWord);
                            if (key.GetValue(Settings.VAL_ENABLE_LOG) == null) 
                                key.SetValue(Settings.VAL_ENABLE_LOG, 0, RegistryValueKind.DWord);
                            if (key.GetValue(Settings.VAL_SPLITTER_POS) == null) 
                                key.SetValue(Settings.VAL_SPLITTER_POS, -1, RegistryValueKind.DWord);
                        }
                    }
                    Log("  Configuration keys created in HKLM", true);
                } catch (Exception ex) { Log("  WARNING: Could not create HKLM config keys: " + ex.Message, true); }

                // 1. .fits -> PerceivedType = image, FullDetails, and InfoTip
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".fits"))
                {
                    key.SetValue("PerceivedType", "image");
                }

                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\SystemFileAssociations\\.fits"))
                {
                    key.SetValue("FullDetails", "prop:System.PropGroup.Image;System.Image.HorizontalSize;System.Image.VerticalSize;System.Image.BitDepth;System.PropGroup.Camera;System.Photo.CameraModel;System.Photo.ExposureTime;System.PropGroup.Description;System.Subject");
                    key.SetValue("InfoTip", "prop:System.ItemType;System.Size;System.Subject;System.Photo.CameraModel;System.Photo.ExposureTime");
                    key.SetValue("PreviewDetails", "prop:*System.Image.HorizontalSize;*System.Image.VerticalSize;*System.Photo.CameraModel;*System.Photo.ExposureTime;*System.Subject");
                }

                // 2. .fits -> ShellEx -> {8895b1c6-b41f-4c1c-a562-0d564250836f} = Preview Handler CLSID
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".fits\\ShellEx\\{8895b1c6-b41f-4c1c-a562-0d564250836f}"))
                {
                    key.SetValue("", guid);
                }

                // 2b. .fits -> ShellEx -> {BB2E617C-0920-11d1-9A0B-00C04FC2D6C1} = Property Handler CLSID
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".fits\\ShellEx\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}"))
                {
                    key.SetValue("", guid);
                }
                
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\SystemFileAssociations\\.fits\\ShellEx\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}"))
                {
                    key.SetValue("", guid);
                }

                // 2c. .fits -> ShellEx -> {e357fccd-...} = Thumbnail Provider CLSID
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".fits\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}"))
                {
                    key.SetValue("", guid);
                }
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Classes\\SystemFileAssociations\\.fits\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}"))
                {
                    key.SetValue("", guid);
                }

                // 3. Register in the official Preview Handlers list
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PreviewHandlers"))
                {
                    key.SetValue(guid, "Fits Preview Handler");
                }

                // 3b. Register in the official Property Handlers list
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\.fits"))
                {
                    key.SetValue("", guid);
                }

                // 4. Set AppID for our CLSID (referencing the standard System surrogate)
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\" + guid))
                {
                    key.SetValue("AppID", appid);
                    key.DeleteValue("ManualSafeSave", false);
                    key.DeleteValue("DisableProcessIsolation", false);
                }
                
                Log("  Registration complete OK", true);
            }
            catch (Exception ex)
            {
                Log("  ERROR during registration: " + ex, true);
                throw; 
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            try
            {
                string guid = t.GUID.ToString("B").ToUpper();
                Log($"--- Unregistering Preview Handler {guid} ---", true);

                // Remove .fits association
                Registry.ClassesRoot.DeleteSubKeyTree(".fits\\ShellEx\\{8895b1c6-b41f-4c1c-a562-0d564250836f}", false);
                Registry.ClassesRoot.DeleteSubKeyTree(".fits\\ShellEx\\{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}", false);
                Registry.ClassesRoot.DeleteSubKeyTree(".fits\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}", false);

                // Remove from PreviewHandlers list
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PreviewHandlers", true))
                {
                    if (key != null) key.DeleteValue(guid, false);
                }

                // Remove from PropertyHandlers list
                Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\.fits", false);

                // Clean up AppID reference
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey("CLSID\\" + guid, true))
                {
                    if (key != null) key.DeleteValue("AppID", false);
                }

                // We keep the settings in HKCU/AppDataLow as they might be useful if reinstalled,
                // and it's safer not to touch user-specific AppDataLow registry during unregister if not strictly needed.

                Log("  Unregistration complete OK", true);
            }
            catch (Exception ex)
            {
                Log("  ERROR during unregistration: " + ex, true);
            }
        }

        #endregion
    }
}
