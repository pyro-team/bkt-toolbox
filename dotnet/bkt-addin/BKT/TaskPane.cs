/*
 * Created by SharpDevelop.
 * User: rdebeerst
 * 
 * Registered Windows-Forms-UserControl representing the TaskPane-Area for BKT.
 * Integrates a single WPF-UserControl.
 *
 */


using System;
using System.Runtime.InteropServices;

using System.Diagnostics;


namespace BKT
{
    [ComVisible(true)]
    [ProgId("BKT.TaskPane")]
    [Guid("76FD3062-86C8-11E4-BE43-6336340000B1")]
    public class TaskPane : System.Windows.Forms.UserControl 
    {
        internal System.Windows.Forms.Integration.ElementHost _host;
        
        TaskPaneControl usrctrl;
		//private dynamic python_delegate;
        
        public TaskPane()
        {
            usrctrl = new TaskPaneControl();
            _host = new System.Windows.Forms.Integration.ElementHost();
            _host.Child = usrctrl;
            this.Controls.Add(_host);
            
            // Standard Fester-grau: Background=""#D7DCE3""
            // helles TaskPane-grau: Background=""#E4E8EE""
        }
        
        
        public TaskPaneControl WpfControl
        {
            get { return usrctrl; }
            set {  }
        }
        
        
        protected override void OnResize(EventArgs e)
        {
            _host.Height = this.Height;
            _host.Width = this.Width;
        }
        
        
		private void Message(string s) {
			System.Windows.Forms.MessageBox.Show(s);
		}
        
        private void DebugMessage(string message)
        {
            Debug.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + ": " + message);
        }
        
        
    }
}