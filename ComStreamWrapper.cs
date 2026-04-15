using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace FitsPreviewHandler
{
    /// <summary>
    /// A .NET Stream wrapper for the COM IStream interface.
    /// Dispose() releases the COM reference via Marshal.ReleaseComObject so the
    /// shell can free the underlying Win32 file handle promptly.
    /// </summary>
    public class ComStreamWrapper : Stream
    {
        private IStream _source;
        private bool _disposed;
        private bool _leaveOpen;

        // Reuse pointers to avoid thousands of Alloc/Free calls during rendering
        private readonly IntPtr _readPtr = Marshal.AllocCoTaskMem(sizeof(int));
        private readonly IntPtr _seekPtr = Marshal.AllocCoTaskMem(sizeof(long));

        public ComStreamWrapper(IStream source, bool leaveOpen = false)
        {
            _source = source ?? throw new ArgumentNullException(nameof(source));
            _leaveOpen = leaveOpen;
        }

        public override bool CanRead  => !_disposed;
        public override bool CanSeek  => !_disposed;
        public override bool CanWrite => false;

        public override long Length
        {
            get
            {
                if (_disposed) throw new ObjectDisposedException(nameof(ComStreamWrapper));
                System.Runtime.InteropServices.ComTypes.STATSTG stat;
                _source.Stat(out stat, 1); // STATFLAG_NONAME = 1
                return stat.cbSize;
            }
        }

        public override long Position
        {
            get => Seek(0, SeekOrigin.Current);
            set => Seek(value, SeekOrigin.Begin);
        }

        public override void Flush() { }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ComStreamWrapper));
            if (offset != 0) throw new NotSupportedException("Only 0 offset supported for ComStream bridge");

            // Native call to explorer.exe (or whoever owns the stream)
            _source.Read(buffer, count, _readPtr);
            return Marshal.ReadInt32(_readPtr);
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ComStreamWrapper));
            _source.Seek(offset, (int)origin, _seekPtr);
            return Marshal.ReadInt64(_seekPtr); 
        }

        public override void SetLength(long value)  => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
                if (_source != null)
                {
                    if (!_leaveOpen)
                    {
                        try { Marshal.ReleaseComObject(_source); } catch { }
                    }
                    _source = null;
                }
                // Memory cleanup
                Marshal.FreeCoTaskMem(_readPtr);
                Marshal.FreeCoTaskMem(_seekPtr);
            }
            base.Dispose(disposing);
        }
    }
}
