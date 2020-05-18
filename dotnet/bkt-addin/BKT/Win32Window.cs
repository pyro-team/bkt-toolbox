using System;

namespace BKT
{
    public class Win32Window : System.Windows.Forms.IWin32Window
    {
        public Win32Window(int windowHandle)
        {
            _windowHandle = new IntPtr(windowHandle);
        }

        private readonly IntPtr _windowHandle;

        public IntPtr Handle
        {
            get { return _windowHandle; }
        }
    }
}