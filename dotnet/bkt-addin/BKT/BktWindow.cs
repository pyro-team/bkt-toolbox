using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Drawing;

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

        protected AddIn addin;

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
            DebugMessage("Initialized");
            base.OnInitialized(e);

            //Store interop helper for later operations
            _helper = new WindowInteropHelper(this);

            //Set window owner
            IntPtr windowID;
            if (addin != null)
            {
                windowID = new IntPtr(addin.GetWindowHandle());
            } else {
                windowID = Process.GetCurrentProcess().MainWindowHandle;
            }

            //FIXME: Process.MainWindowHandle is not always reliable, better to pass context and use activewindow from app
            // IntPtr windowID = Process.GetCurrentProcess().MainWindowHandle;
            _helper.Owner = windowID;
            DebugMessage(string.Format("Set owner: {0}", windowID));
            // _helper.Owner = GetForegroundWindow();

            // DPIHelper.SetDpiAwareness(DPIHelper.PROCESS_DPI_AWARENESS.Process_Per_Monitor_DPI_Aware);
            DPIHelper.SetThreadDpiAwareness(DPIHelper.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
            DebugMessage("DPI Awareness changed");
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            DebugMessage("Source Initialized");
            base.OnSourceInitialized(e);

            double scaling = GetScalingForWindow(_helper.Owner);

            if (IsPopup || IsToolbar)
            {
                DebugMessage("Apply Popup Style");
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
        //     DebugMessage("Setting owner");
        //     var helper = new WindowInteropHelper(this);
        //     helper.Owner = handle;
        // }

        // public void SetCurrentWindowAsOwner()
        // {
        //     SetOwner(Process.GetCurrentProcess().MainWindowHandle);
        // }

        public void SetOwner(int windowID)
        {
            DebugMessage(string.Format("Set new owner: {0}", windowID));
            _helper.Owner = new IntPtr(windowID);
        }

        public void SetDevicePosition(double? deviceLeft=null, double? deviceTop=null, double? deviceRight=null, double? deviceBottom=null)
        {
            DebugMessage("Setting device position");
            try {
                // var source = HwndSource.FromHwnd(_helper.EnsureHandle());
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
                
                DebugMessage(string.Format("SetDevicePosition x={0}, y={1}", x, y));

                // var p = this.PointToScreen(new System.Windows.Point(x, y));
                // DebugMessage(string.Format("PointToScreen x={0}, y={1}", p.X, p.Y));

                double VsTop = SystemParameters.VirtualScreenTop;
                double VsLeft = SystemParameters.VirtualScreenLeft;
                double VsRight = VsLeft + SystemParameters.VirtualScreenWidth;
                double VsBottom = VsTop + SystemParameters.VirtualScreenHeight;

                DebugMessage(string.Format("VirtualScreen top {0}, left {1}, right {2}, bottom {3}", VsTop, VsLeft, VsRight, VsBottom));

                // var transform = source.CompositionTarget.TransformFromDevice;
                // var ptLogicalUnits = transform.Transform(new Point(x, y));
                // if (deviceLeft != null) {
                //     this.Left = Math.Min(Math.Max(ptLogicalUnits.X, VsLeft), VsRight - this.Width);
                // } else if (deviceRight != null) {
                //     this.Left = Math.Min(Math.Max(ptLogicalUnits.X - this.Width, VsLeft), VsRight - this.Width);
                // }
                // if (deviceTop != null) {
                //     this.Top  = Math.Min(Math.Max(ptLogicalUnits.Y, VsTop), VsBottom - this.Height);
                // } else if (deviceBottom != null) {
                //     this.Top  = Math.Min(Math.Max(ptLogicalUnits.Y - this.Height, VsTop), VsBottom - this.Height);
                // }

                double scaling = GetScalingForWindow(_helper.Owner);
                DebugMessage(string.Format("Scaling x={0}, y={1}", x*scaling, y*scaling));

                // this.TranslatePoint()
                // var p = this.PointToScreen(new System.Windows.Point(x,y));
                // x = p.X;
                // y = p.Y;
                // double scaling = 0.5;

                if (deviceLeft != null) {
                    this.Left = Math.Min(Math.Max(x*scaling, VsLeft), VsRight - this.Width);
                    // this.Left = x;
                } else if (deviceRight != null) {
                    this.Left = Math.Min(Math.Max(x*scaling - this.Width, VsLeft), VsRight - this.Width);
                }
                if (deviceTop != null) {
                    this.Top  = Math.Min(Math.Max(y*scaling, VsTop), VsBottom - this.Height);
                } else if (deviceBottom != null) {
                    this.Top  = Math.Min(Math.Max(y*scaling - this.Height, VsTop), VsBottom - this.Height);
                }

                // MoveWindow(_helper.EnsureHandle(), (int) x, (int) y, (int) this.Width, (int) this.Height, true);

            } catch (Exception e) {
                DebugMessage(e.ToString());
            }
        }

        public void SetDeviceSize(double? deviceWidth=null, double? deviceHeight=null)
        {
            DebugMessage("Setting device size");
            try {
                // var source = HwndSource.FromHwnd(_helper.EnsureHandle());
                // var transform = source.CompositionTarget.TransformFromDevice;
                // if (deviceWidth != null) {
                //     this.Width = Math.Min(SystemParameters.VirtualScreenWidth, transform.M11 * (double)deviceWidth);
                // }
                // if (deviceHeight != null) {
                //     this.Height = Math.Min(SystemParameters.VirtualScreenHeight, transform.M22 * (double)deviceHeight);
                // }

                double scaling = GetScalingForWindow(_helper.Owner);
                if (deviceWidth != null) {
                    this.Width = Math.Min(SystemParameters.VirtualScreenWidth, scaling * (double)deviceWidth);
                }
                if (deviceHeight != null) {
                    this.Height = Math.Min(SystemParameters.VirtualScreenHeight, scaling * (double)deviceHeight);
                }
            } catch (Exception e) {
                DebugMessage(e.ToString());
            }
        }

        public System.Windows.Point GetTransformFrom(double x, double y)
        {
                // var source = HwndSource.FromHwnd(_helper.EnsureHandle());
                var source = PresentationSource.FromVisual(this);
                var transform = source.CompositionTarget.TransformFromDevice;
                DebugMessage(string.Format("GetTransformFrom: transform.M11 is {0}", transform.M11));
                DebugMessage(string.Format("GetTransformFrom: transform.M22 is {0}", transform.M22));
                return transform.Transform(new System.Windows.Point(x, y));
        }

        public System.Windows.Point GetTransformTo(double x, double y)
        {
                // var source = HwndSource.FromHwnd(_helper.EnsureHandle());
                var source = PresentationSource.FromVisual(this);
                var transform = source.CompositionTarget.TransformToDevice;
                DebugMessage(string.Format("GetTransformTo: transform.M11 is {0}", transform.M11));
                DebugMessage(string.Format("GetTransformTo: transform.M22 is {0}", transform.M22));
                return transform.Transform(new System.Windows.Point(x, y));
        }

        public void ApplyScaleTransform(uint dpi)
        {
            DebugMessage("ApplyScaleTransform");
            var source = PresentationSource.FromVisual(this);
            // if (source == null) return;
            var wpfDpi = 96 * source.CompositionTarget.TransformToDevice.M11;
            var scaleFactor = dpi / wpfDpi;
            System.Windows.Media.VisualTreeHelper.SetRootDpi(this, new DpiScale(scaleFactor,scaleFactor));
            DebugMessage(string.Format("ApplyScaleTransform: scaleFactor is {0}", scaleFactor));
            var scaleTransform = 
                Math.Abs(scaleFactor - 1.0) < 0.001 ? null : new System.Windows.Media.ScaleTransform(scaleFactor, scaleFactor);
            this.SetValue(FrameworkElement.LayoutTransformProperty, scaleTransform);
            // this.LayoutTransform = scaleTransform;
            this.OnDpiChanged(new DpiScale(1.0, 1.0), new DpiScale(scaleFactor,scaleFactor));
        }

        public double GetScalingForWindow(IntPtr hWnd)
        {
            try {
#if DEBUG
                // var source = HwndSource.FromHwnd(_helper.EnsureHandle());
                var source = PresentationSource.FromVisual(this);
                var transform = source.CompositionTarget.TransformFromDevice;
                DebugMessage(string.Format("GetWindowDpi: transform.M11 is {0}", transform.M11));
                DebugMessage(string.Format("GetWindowDpi: transform.M22 is {0}", transform.M22));
                DebugMessage(string.Format("GetWindowDpi: transform.OffsetX is {0}", transform.OffsetX));
                DebugMessage(string.Format("GetWindowDpi: transform.OffsetY is {0}", transform.OffsetY));

                Graphics g = Graphics.FromHwnd(IntPtr.Zero);
                DebugMessage(string.Format("GetWindowDpi: g.DpiX is {0}", g.DpiX));
                DebugMessage(string.Format("GetWindowDpi: g.DpiY is {0}", g.DpiY));

                var dpiscale = System.Windows.Media.VisualTreeHelper.GetDpi(this);
                DebugMessage(string.Format("VisualTreeHelper: DpiScaleX is {0}", dpiscale.DpiScaleX));
                DebugMessage(string.Format("VisualTreeHelper: DpiScaleY is {0}", dpiscale.DpiScaleY));
                DebugMessage(string.Format("VisualTreeHelper: PixelsPerDip is {0}", dpiscale.PixelsPerDip));
                DebugMessage(string.Format("VisualTreeHelper: PixelsPerInchX is {0}", dpiscale.PixelsPerInchX));
                DebugMessage(string.Format("VisualTreeHelper: PixelsPerInchY is {0}", dpiscale.PixelsPerInchY));

                DebugMessage(DPIHelper.GetThreadDpiAwareness().ToString());
                DebugMessage(DPIHelper.GetProcessDpiAwareness().ToString());
                DebugMessage(DPIHelper.GetWindowDpiAwareness(_helper.EnsureHandle()).ToString());
#endif

                uint dpi = GetDpiForWindow(_helper.Owner);
                DebugMessage(string.Format("GetWindowDpi: DPI is {0}, scaling {1}, for window {2}", dpi, 96.0 / dpi, _helper.Owner));
                ApplyScaleTransform(dpi);
                return 96.0 / dpi;
            } catch (Exception e) {
                DebugMessage(e.ToString());
                return 1.0;
            }
        }

        public Rect GetDeviceRect()
        {
            DebugMessage("Getting device rect");
            try {
                // var source = HwndSource.FromHwnd(_helper.EnsureHandle());
                // var ltPhysicalUnits = source.CompositionTarget.TransformToDevice.Transform(new Point(this.Left, this.Top));
                // var brPhysicalUnits = source.CompositionTarget.TransformToDevice.Transform(new Point(this.Left+this.Width, this.Top+this.Height));
                // return new Rect(ltPhysicalUnits, brPhysicalUnits);

                double scaling = GetScalingForWindow(_helper.Owner);
                return new Rect(this.Left*scaling, this.Top*scaling, this.Width*scaling, this.Height*scaling);
            } catch (Exception e) {
                DebugMessage(e.ToString());
                return new Rect(0,0,0,0);
            }
        }

        public void ShiftWindowOntoScreen()
        {
            DebugMessage("Shift window back onto screen");
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
            Debug.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + ": [BKT Window " + this.Title + "] " + message);
        }

        [DllImport("user32.dll")]
        private static extern uint SetWindowLong(IntPtr hWnd, int nIndex, uint dwNewLong);

        [DllImport("user32.dll")]
        private static extern uint GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport( "user32.dll", SetLastError = true )]
        private static extern bool MoveWindow( IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint );

        private const int GWL_STYLE = (-16);
        private const int GWL_EXSTYLE = (-20);
        private const uint WS_POPUP = 0x80000000;
        private const uint WS_EX_NOACTIVATE = 0x08000000;
        private const uint WS_EX_TOOLWINDOW = 0x00000080;
        private const int WM_MOUSEACTIVATE = 0x0021;
        private const int MA_NOACTIVATE = 0x0003;

    }
}