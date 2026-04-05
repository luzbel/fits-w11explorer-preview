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
    
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("fc4801a3-2ba9-11cf-a229-00aa003d7352")]
    public interface IObjectWithSite
    {
        void SetSite([In, MarshalAs(UnmanagedType.IUnknown)] object pUnkSite);
        void GetSite(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvSite);
    }

    [ComVisible(true)]
    [Guid("AF1C3D6A-81E9-4F5B-9A8C-2D9E71F04B3E")]
    [ClassInterface(ClassInterfaceType.None)]
    public class FitsPreviewHandlerExtension : IPreviewHandler, IInitializeWithFile, IInitializeWithStream, IInitializeWithItem, IObjectWithSite
    {
        private string _filePath;
        private IntPtr _parentHwnd;
        private RECT _bounds;
        private FitsPreviewControl _control;
        private System.Threading.Thread _uiThread;
        private Stream _stream; // Direct zero-copy stream
        
        // Shared log writer — delegates to the same path used by FitsPreviewControl
        private static void Log(string msg)
        {
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
        }

        void IInitializeWithStream.Initialize(object pstream, uint grfMode)
        {
            Log($"IInitializeWithStream.Initialize — grfMode={grfMode}, stream is {pstream?.GetType().FullName ?? "null"}");
            try
            {
                if (pstream is System.Runtime.InteropServices.ComTypes.IStream comStream)
                {
                    _stream = new ComStreamWrapper(comStream);
                    Log("  Stream wrapped OK — ready for Zero-Copy");
                }
                else
                {
                    Log("  WARNING: pstream does not implement IStream");
                }
            }
            catch (Exception ex) { Log("  EXCEPTION in IInitializeWithStream.Initialize: " + ex); }
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
                _control.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        _control.Dispose();
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

        #region Registration Logic

        [ComRegisterFunction]
        public static void Register(Type t)
        {
            try
            {
                string guid = t.GUID.ToString("B").ToUpper();
                // Standard Microsoft AppID for Preview Host (prevhost.exe)
                string appid = "{6d2b5079-2f0b-48dd-ab7f-97cec514d30b}";
                
                Log($"--- Registering Preview Handler {guid} ---");

                // 1. .fits -> PerceivedType = image
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".fits"))
                {
                    key.SetValue("PerceivedType", "image");
                }

                // 2. .fits -> ShellEx -> {8895b1c6-b41f-4c1c-a562-0d564250836f} = our CLSID
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".fits\\ShellEx\\{8895b1c6-b41f-4c1c-a562-0d564250836f}"))
                {
                    key.SetValue("", guid);
                }

                // 3. Register in the official Preview Handlers list
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PreviewHandlers"))
                {
                    key.SetValue(guid, "Fits Preview Handler");
                }

                // 4. Set AppID for our CLSID (referencing the standard System surrogate)
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\" + guid))
                {
                    key.SetValue("AppID", appid);
                }

                // NOTE: We no longer need to manually set InprocServer32/CodeBase 
                // because regasm /codebase handles it when used correctly.
                
                Log("  Registration complete OK");
            }
            catch (Exception ex)
            {
                Log("  ERROR during registration: " + ex);
                throw; 
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            try
            {
                string guid = t.GUID.ToString("B").ToUpper();
                Log($"--- Unregistering Preview Handler {guid} ---");

                // Remove .fits association
                Registry.ClassesRoot.DeleteSubKeyTree(".fits\\ShellEx\\{8895b1c6-b41f-4c1c-a562-0d564250836f}", false);

                // Remove from PreviewHandlers list
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PreviewHandlers", true))
                {
                    if (key != null) key.DeleteValue(guid, false);
                }

                // Clean up AppID reference (optional, regasm /u cleans most CLSID keys)
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey("CLSID\\" + guid, true))
                {
                    if (key != null) {
                        key.DeleteValue("AppID", false);
                    }
                }

                Log("  Unregistration complete OK");
            }
            catch (Exception ex)
            {
                Log("  ERROR during unregistration: " + ex);
            }
        }

        #endregion
    }
}
