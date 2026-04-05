using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace FitsPreviewHandler
{
    /// <summary>
    /// A .NET Stream wrapper for the COM IStream interface.
    /// This allows us to read FITS data directly from the Shell without copying to a disk file.
    /// </summary>
    public class ComStreamWrapper : Stream
    {
        private readonly IStream _source;

        public ComStreamWrapper(IStream source)
        {
            _source = source ?? throw new ArgumentNullException(nameof(source));
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;

        public override long Length
        {
            get
            {
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
            // Note: IStream.Read starts at the current pointer position.
            // We assume offset is 0 for simplicity in this bridge.
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

        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            // We do NOT dispose the _source IStream here because the Windows Shell owns its lifecycle.
            // The shell will release the COM object when the preview is unloaded.
            base.Dispose(disposing);
        }
    }
}
