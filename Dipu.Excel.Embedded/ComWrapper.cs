using System;
using System.Runtime.InteropServices;

namespace Dipu.Excel.Embedded
{
    public static class Com
    {
        public static ComWrapper<T> Wrap<T>(T comObject) where T : class => new ComWrapper<T>(comObject);
    }

    public class ComWrapper<T> : IDisposable where T : class
    {
        public readonly T ComObject;

        public ComWrapper(T comObject)
        {
            ComObject = comObject;
        }

        public override bool Equals(object obj)
        {
            if (obj is ComWrapper<T> that)
            {
                return this.ComObject.Equals(that.ComObject);
            }

            return this == obj;
        }

        public override int GetHashCode()
        {
            return this.ComObject.GetHashCode();
        }

        ~ComWrapper()
        {
            ReleaseUnmanagedResources();
        }

        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }

        private void ReleaseUnmanagedResources()
        {
            Marshal.ReleaseComObject(this.ComObject);
        }
    }
}
