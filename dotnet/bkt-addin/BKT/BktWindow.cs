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
        public static readonly DependencyProperty IsToolbarProperty = DependencyProperty.Register("IsToolbar", typeof(bool), typeof(BktWindow), new FrameworkPropertyMetadata(false));

        private WindowInteropHelper _helper;

        //TODO: Make this property available in XAML, but how?
        public bool IsPopup
        {
            get { return (bool)GetValue(IsPopupProperty); }
            set { SetValue(IsPopupProperty, value); }
        }
        public bool IsToolbar
        {
            get { return (bool)GetValue(IsToolbarProperty); }
            set { SetValue(IsToolbarProperty, value); }
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

            if (IsPopup || IsToolbar)
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

        protected override void OnMouseRightButtonUp(System.Windows.Input.MouseButtonEventArgs e)
        {
            if (IsPopup)
            {
                this.Close();
            }
            base.OnMouseRightButtonUp(e);
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

        public void SetDevicePosition(double? deviceLeft=null, double? deviceTop=null, double? deviceRight=null, double? deviceBottom=null)
        {
            DebugMessage("BKT Window: Setting device position");
            try {
                var source = HwndSource.FromHwnd(_helper.EnsureHandle());
                // var source = PresentationSource.FromVisual(this);
                double x = 0;
                double y = 0;
                if (deviceLeft != null) {
                    x = (double)deviceLeft;
                } else if (deviceRight != null) {
                    x = (double)deviceRight;
                }
                if (deviceTop != null) {
                    y = (double)deviceTop;
                } else if (deviceBottom != null) {
                    y = (double)deviceBottom;
                }

                double VsTop = SystemParameters.VirtualScreenTop;
                double VsLeft = SystemParameters.VirtualScreenLeft;
                double VsRight = VsLeft + SystemParameters.VirtualScreenWidth;
                double VsBottom = VsTop + SystemParameters.VirtualScreenHeight;

                var ptLogicalUnits = source.CompositionTarget.TransformFromDevice.Transform(new Point(x, y));
                if (deviceLeft != null) {
                    this.Left = Math.Min(Math.Max(ptLogicalUnits.X, VsLeft), VsRight - this.Width);
                } else if (deviceRight != null) {
                    this.Left = Math.Min(Math.Max(ptLogicalUnits.X - this.Width, VsLeft), VsRight - this.Width);
                }
                if (deviceTop != null) {
                    this.Top  = Math.Min(Math.Max(ptLogicalUnits.Y, VsTop), VsBottom - this.Height);
                } else if (deviceBottom != null) {
                    this.Top  = Math.Min(Math.Max(ptLogicalUnits.Y - this.Height, VsTop), VsBottom - this.Height);
                }
            } catch (Exception e) {
                DebugMessage(e.ToString());
            }
        }

        public void SetDeviceSize(double? deviceWidth=null, double? deviceHeight=null)
        {
            DebugMessage("BKT Window: Setting device size");
            try {
                var source = HwndSource.FromHwnd(_helper.EnsureHandle());
                var transform = source.CompositionTarget.TransformFromDevice;
                if (deviceWidth != null) {
                    this.Width = Math.Min(SystemParameters.VirtualScreenWidth, transform.M11 * (double)deviceWidth);
                }
                if (deviceHeight != null) {
                    this.Height = Math.Min(SystemParameters.VirtualScreenHeight, transform.M22 * (double)deviceHeight);
                }
            } catch (Exception e) {
                DebugMessage(e.ToString());
            }
        }

        public Rect GetDeviceRect()
        {
            DebugMessage("BKT Window: Getting device rect");
            try {
                var source = HwndSource.FromHwnd(_helper.EnsureHandle());
                var ltPhysicalUnits = source.CompositionTarget.TransformToDevice.Transform(new Point(this.Left, this.Top));
                var brPhysicalUnits = source.CompositionTarget.TransformToDevice.Transform(new Point(this.Left+this.Width, this.Top+this.Height));
                return new Rect(ltPhysicalUnits, brPhysicalUnits);
            } catch (Exception e) {
                DebugMessage(e.ToString());
                return new Rect(0,0,0,0);
            }
        }

        public void ShiftWindowOntoScreen()
        {
            DebugMessage("BKT Window: Shift window back onto screen");
            double VsTop = SystemParameters.VirtualScreenTop;
            double VsLeft = SystemParameters.VirtualScreenLeft;
            double VsRight = VsLeft + SystemParameters.VirtualScreenWidth;
            double VsBottom = VsTop + SystemParameters.VirtualScreenHeight;

            if (this.Top < VsTop) {
                this.Top = VsTop;
            } else if (this.Top + this.Height > VsBottom) {
                this.Top = VsBottom - this.Height;
            }

            if (this.Left < VsLeft) {
                this.Left = VsLeft;
            } else if (this.Left + this.Width > VsRight) {
                this.Left = VsRight - this.Width;
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