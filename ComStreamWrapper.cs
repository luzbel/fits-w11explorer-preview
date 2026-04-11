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

        public ComStreamWrapper(IStream source)
        {
            _source = source ?? throw new ArgumentNullException(nameof(source));
        }

        public override bool CanRead  => !_disposed;
        public override bool CanSeek  => !_disposed;
        public override bool CanWrite => false;

        public override long Length
        {
            get
            {
                if (_disposed) throw new ObjectDisposedException(nameof(ComStreamWrapper));
                _source.Stat(out System.Runtime.InteropServices.ComTypes.STATSTG stat, 1); // STATFLAG_NONAME = 1
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

            IntPtr bytesReadPtr = Marshal.AllocCoTaskMem(sizeof(int));
            try
            {
                _source.Read(buffer, count, bytesReadPtr);
                return Marshal.ReadInt32(bytesReadPtr);
            }
            finally
            {
                Marshal.FreeCoTaskMem(bytesReadPtr);
            }
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ComStreamWrapper));
            IntPtr posPtr = Marshal.AllocCoTaskMem(sizeof(long));
            try
            {
                _source.Seek(offset, (int)origin, posPtr);
                return Marshal.ReadInt64(posPtr);
            }
            finally
            {
                Marshal.FreeCoTaskMem(posPtr);
            }
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
                    // Release our COM reference so the shell's IStream ref-count can reach
                    // zero and the underlying Win32 file handle can be closed immediately.
                    try { Marshal.ReleaseComObject(_source); } catch { }
                    _source = null;
                }
            }
            base.Dispose(disposing);
        }
    }
}
