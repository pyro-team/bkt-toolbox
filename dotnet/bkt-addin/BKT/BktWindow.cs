using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

/// <summary>
/// Popup with code to not be the topmost control
/// </summary>


namespace BKT
{
    public class BktWindow : Window
    {
        public static readonly DependencyProperty IsPopupProperty = DependencyProperty.Register("IsPopup", typeof(bool), typeof(BktWindow), new FrameworkPropertyMetadata(false));

        private WindowInteropHelper _helper;

        //TODO: Make this property available in XAML, but how?
        public bool IsPopup
        {
            get { return (bool)GetValue(IsPopupProperty); }
            set { SetValue(IsPopupProperty, value); }
        }

        protected override void OnInitialized(EventArgs e)
        {
            DebugMessage("BKT Window: Initialized");
            base.OnInitialized(e);

            //Store interop helper for later operations
            _helper = new WindowInteropHelper(this);

            //Set window owner
            //FIXME: Process.MainWindowHandle is not always reliable, better to pass context and use activewindow from app
            _helper.Owner = Process.GetCurrentProcess().MainWindowHandle;
            // _helper.Owner = GetForegroundWindow();
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            DebugMessage("BKT Window: Source Initialized");
            base.OnSourceInitialized(e);

            if (IsPopup)
            {
                DebugMessage("BKT Window: Apply Popup Style");
                //Never activate window
                var source = HwndSource.FromHwnd(_helper.EnsureHandle());
                // var source = PresentationSource.FromVisual(this) as HwndSource;
                source.AddHook(WndProc);
                //Set popup style
                SetWindowLong(_helper.Handle, GWL_STYLE,
                    GetWindowLong(_helper.Handle, GWL_STYLE) | WS_POPUP);
                //Set no activate extended style
                SetWindowLong(_helper.Handle, GWL_EXSTYLE,
                    GetWindowLong(_helper.Handle, GWL_EXSTYLE) | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE);
            }
        }

        // public void SetOwner(IntPtr handle)
        // {
        //     DebugMessage("BKT Window: Setting owner");
        //     var helper = new WindowInteropHelper(this);
        //     helper.Owner = handle;
        // }

        // public void SetCurrentWindowAsOwner()
        // {
        //     SetOwner(Process.GetCurrentProcess().MainWindowHandle);
        // }

        public void SetOwner(int windowID)
        {
            DebugMessage(string.Format("BKT Window: Set new owner: {0}", windowID));
            _helper.Owner = new IntPtr(windowID);
        }

        public void SetDevicePosition(int deviceLeft, int deviceTop)
        {
            DebugMessage("BKT Window: Setting device position");
            try {
                var source = HwndSource.FromHwnd(_helper.EnsureHandle());
                // var source = PresentationSource.FromVisual(this);
                var ptLogicalUnits = source.CompositionTarget.TransformFromDevice.Transform(new Point(deviceLeft, deviceTop));
                this.Left = ptLogicalUnits.X;
                this.Top  = ptLogicalUnits.Y;
            } catch (Exception e) {
                DebugMessage(e.ToString());
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_MOUSEACTIVATE)
            {
                handled = true;
                return new IntPtr(MA_NOACTIVATE);
            }
            else
            {
                return IntPtr.Zero;
            }
        }

        private void DebugMessage(string message)
        {
            Debug.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + ": " + message);
        }

        [DllImport("user32.dll")]
        private static extern uint SetWindowLong(IntPtr hWnd, int nIndex, uint dwNewLong);

        [DllImport("user32.dll")]
        private static extern uint GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private const int GWL_STYLE = (-16);
        private const int GWL_EXSTYLE = (-20);
        private const uint WS_POPUP = 0x80000000;
        private const uint WS_EX_NOACTIVATE = 0x08000000;
        private const uint WS_EX_TOOLWINDOW = 0x00000080;
        private const int WM_MOUSEACTIVATE = 0x0021;
        private const int MA_NOACTIVATE = 0x0003;

    }
}