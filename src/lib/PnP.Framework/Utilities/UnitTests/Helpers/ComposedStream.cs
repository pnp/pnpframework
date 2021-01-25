using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PnP.Framework.Utilities.UnitTests.Helpers
{
    public class ComposedStream : Stream
    {
        public override bool CanRead { get; } = true;

        public override bool CanSeek { get; } = true;

        public override bool CanWrite { get; } = true;

        public override long Length
        {
            get
            {
                return BaseStream.Length;
            }
        }
        public Stream BaseStream { get; }
        public ComposedStream(Stream baseStream)
        {
            BaseStream = baseStream;
        }

        public override long Position
        {
            get
            {
                return BaseStream.Position;
            }
            set
            {
                BaseStream.Position = value;
            }
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return BaseStream.Read(buffer, offset, count);
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return BaseStream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            BaseStream.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            BaseStream.Write(buffer, offset, count);
        }
        public override void Close()
        {
            base.Close();
        }
    }
}
