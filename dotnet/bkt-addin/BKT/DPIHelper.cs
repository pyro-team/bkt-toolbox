using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace BKT
{
    /// <summary>
    /// https://github.com/shawnmcdowell/Office-Extensibility-Samples/tree/master/VSTO%20SharedAddin/SharedModule
    /// </summary>
    internal static class DPIHelper
    {
        // Hack until sdk version of windef.h is updated
        public enum DPI_HOSTING_BEHAVIOR
        {
            DPI_HOSTING_BEHAVIOR_INVALID = -1,
            DPI_HOSTING_BEHAVIOR_DEFAULT = 0,
            DPI_HOSTING_BEHAVIOR_MIXED = 1
        }

        [DllImport("SHCore.dll", SetLastError = true)]
        private static extern bool SetProcessDpiAwareness(PROCESS_DPI_AWARENESS awareness);

        [DllImport("SHCore.dll", SetLastError = true)]
        private static extern void GetProcessDpiAwareness(IntPtr hprocess, out PROCESS_DPI_AWARENESS awareness);

        /// <summary>
        /// https://docs.microsoft.com/en-us/windows/desktop/api/Winuser/nf-winuser-setthreaddpiawarenesscontext
        /// </summary>
        /// <param name="dpiContext"></param>
        /// <returns></returns>
        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_AWARENESS_CONTEXT SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT awareness);

        /// <summary>
        /// https://docs.microsoft.com/en-us/windows/desktop/api/Winuser/nf-winuser-getthreaddpiawarenesscontext
        /// </summary>
        /// <returns></returns>
        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_AWARENESS_CONTEXT GetThreadDpiAwarenessContext();

        /// <summary>
        /// https://docs.microsoft.com/en-us/windows/desktop/api/Winuser/nf-winuser-getwindowdpiawarenesscontext
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_AWARENESS_CONTEXT GetWindowDpiAwarenessContext(IntPtr hWnd);

        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_AWARENESS GetAwarenessFromDpiAwarenessContext(DPI_AWARENESS_CONTEXT value);

        // DPI_HOSTING_BEHAVIOR WINAPI SetThreadDpiHostingBehavior(_In_ DPI_HOSTING_BEHAVIOR dpiHostingBehavior);
        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_HOSTING_BEHAVIOR SetThreadDpiHostingBehavior(DPI_HOSTING_BEHAVIOR dpiHostingBehavior);

        // DPI_HOSTING_BEHAVIOR WINAPI GetThreadDpiHostingBehavior(_In_ HWND hwnd);
        [DllImport("User32.dll", SetLastError = true)]
        private static extern DPI_HOSTING_BEHAVIOR GetThreadDpiHostingBehavior(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, ref Rectangle rect);

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder className, int charCount);

        public enum PROCESS_DPI_AWARENESS
        {
            Process_DPI_Unaware = 0,
            Process_System_DPI_Aware = 1,
            Process_Per_Monitor_DPI_Aware = 2
        }

        /// <summary>
        /// DPI_AWARENESS_CONTEXT; https://docs.microsoft.com/en-us/windows/win32/hidpi/dpi-awareness-context
        /// </summary>
        public struct DPI_AWARENESS_CONTEXT
        {
            private IntPtr value;

            private DPI_AWARENESS_CONTEXT(IntPtr value)
            {
                this.value = value;
            }

            public static implicit operator DPI_AWARENESS_CONTEXT(IntPtr value)
            {
                return new DPI_AWARENESS_CONTEXT(value);
            }

            public static implicit operator IntPtr(DPI_AWARENESS_CONTEXT context)
            {
                return context.value;
            }

            public static bool operator ==(IntPtr context1, DPI_AWARENESS_CONTEXT context2)
            {
                return AreDpiAwarenessContextsEqual(context1, context2);
            }

            public static bool operator !=(IntPtr context1, DPI_AWARENESS_CONTEXT context2)
            {
                return !AreDpiAwarenessContextsEqual(context1, context2);
            }

            public override bool Equals(object obj)
            {
                return base.Equals(obj);
            }

            public override int GetHashCode()
            {
                return base.GetHashCode();
            }
        }

        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_INVALID = IntPtr.Zero;
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_UNAWARE = new IntPtr(-1);
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = new IntPtr(-2);
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE = new IntPtr(-3);
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);

        /// <summary>
        /// DPI_AWARENESS; https://docs.microsoft.com/en-us/windows/desktop/api/windef/ne-windef-dpi_awareness
        /// </summary>
        public enum DPI_AWARENESS
        {
            DPI_AWARENESS_INVALID = -1,
            DPI_AWARENESS_UNAWARE = 0,
            DPI_AWARENESS_SYSTEM_AWARE = 1,
            DPI_AWARENESS_PER_MONITOR_AWARE = 2
        }

        /// <summary>
        /// https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-aredpiawarenesscontextsequal
        /// </summary>
        /// <param name="dpiAwarenessContext1"></param>
        /// <param name="dpiAwarenessContext2"></param>
        /// <returns></returns>
        [DllImport("User32.dll")]
        public static extern bool AreDpiAwarenessContextsEqual(DPI_AWARENESS_CONTEXT dpiAwarenessContext1, DPI_AWARENESS_CONTEXT dpiAwarenessContext2);

        public static bool SetDpiAwareness(PROCESS_DPI_AWARENESS awareness)
        {
            return SetProcessDpiAwareness(awareness);
        }

        public static DPI_AWARENESS_CONTEXT SetThreadDpiAwareness(DPI_AWARENESS_CONTEXT awareness)
        {
            return SetThreadDpiAwarenessContext(awareness);
        }

        public static PROCESS_DPI_AWARENESS GetProcessDpiAwareness()
        {
            return GetProcessDpiAwareness(Process.GetCurrentProcess().Handle);
        }

        public static PROCESS_DPI_AWARENESS GetProcessDpiAwareness(IntPtr hprocess)
        {
            PROCESS_DPI_AWARENESS awareness;
            GetProcessDpiAwareness(hprocess, out awareness);
            return awareness;
        }

        public static DPI_AWARENESS GetThreadDpiAwareness()
        {
            DPI_AWARENESS_CONTEXT context = GetThreadDpiAwarenessContext();
            return GetAwarenessFromDpiAwarenessContext(context);
        }

        public static DPI_AWARENESS GetWindowDpiAwareness(IntPtr hWnd)
        {
            DPI_AWARENESS_CONTEXT context = GetWindowDpiAwarenessContext(hWnd);
            return GetAwarenessFromDpiAwarenessContext(context);
        }

        public static DPI_HOSTING_BEHAVIOR SetChildWindowMixedMode(DPI_HOSTING_BEHAVIOR value)
        {
            return SetThreadDpiHostingBehavior(value);
        }

        public static DPI_HOSTING_BEHAVIOR GetChildWindowMixedMode(IntPtr hWnd)
        {
            return GetThreadDpiHostingBehavior(hWnd);
        }

        public static void DebugPrintDPIAwareness(IntPtr hprocess, string message)
        {
            Debug.WriteLine(DPIAwarenessText(hprocess, message));
        }

        public static string DPIAwarenessText(IntPtr hprocess, string message)
        {
            return String.Format("***{0}: Process {1}, Thread {2}", message, GetProcessDpiAwareness(hprocess), GetThreadDpiAwareness());
        }

        public static string DPIAwarenessText(string message)
        {
            return DPIAwarenessText(Process.GetCurrentProcess().Handle, message);
        }

        public static Rectangle GetWindowRectangle(IntPtr hWnd)
        {
            Rectangle rect = Rectangle.Empty;
            GetWindowRect(hWnd, ref rect);

            return rect;
        }

        public static IntPtr GetParentWindow(IntPtr hWnd)
        {
            return GetParent(hWnd);
        }

        public static string GetWindowClassName(IntPtr hWnd)
        {
            StringBuilder buff = new StringBuilder(256);
            int retCount = 0;

            retCount = GetClassName(hWnd, buff, 256);

            return buff.ToString();
        }

        public static IntPtr FindParentWithClassName(IntPtr hWndChild, string className)
        {
            IntPtr hwndParent = GetParent(hWndChild);

            while (hwndParent != IntPtr.Zero)
            {
                if (GetWindowClassName(hwndParent).Equals(className, StringComparison.InvariantCultureIgnoreCase))
                {
                    return hwndParent;
                }
                hwndParent = GetParent(hwndParent);
            }
            return IntPtr.Zero;
        }
    }
}
