using System;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
//using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Extensibility;

namespace BKT.Dev
{
    /// <summary>
    /// Description of MyClass.
    /// </summary>
    [ProgId("BKT.Dev.DevAddIn")]
	[Guid("FC4DBFDD-A8A2-4675-A32D-A56337844DC4")]
	[ComVisible(true)]
	public class DevAddIn : IDTExtensibility2, IRibbonExtensibility
	{
		Application app;
		
		private void Show(string s) {
			System.Windows.Forms.MessageBox.Show(s);
		}
		
		public void OnConnection(object app, ext_ConnectMode connect_mode, object addin_inst, ref Array custom)
		{
			this.app = (Application) app;
		}
		
		public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
			this.app = null;
		}
		
		public void OnAddInsUpdate(ref Array custom)
		{
		}
		
		public void OnStartupComplete(ref Array custom)
		{
		}
		
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		public string GetCustomUI(string RibbonID)
		{
			Stream input = GetType().Assembly.GetManifestResourceStream("UI_XML");
			StreamReader reader = new StreamReader(input, UTF8Encoding.UTF8);
			string xml_data = reader.ReadToEnd();
			input.Close();
			//Show(xml_data);
			return xml_data;
		}
		
		private COMAddIn GetBKT() {
			if(app == null) {
				return null;
			}
			COMAddIn bkt = null;
			try {
				bkt = app.COMAddIns.Item("BKT.AddIn");
				return bkt;
			} catch (Exception) {
				return null;
			}
		}

		public void ReloadBKTAddIn(IRibbonControl control)
        {
			COMAddIn bkt = GetBKT();
			if(bkt == null) {
				Show("BKT not found");
				return;
			}
			if(bkt.Connect) {
				bkt.Connect = false;
				bkt.Connect = true;
			} else {
				bkt.Connect = true;
			}
        }
		
		public void UnloadBKTAddIn(IRibbonControl control)
        {
			COMAddIn bkt = GetBKT();
			if(bkt == null) {
				Show("BKT not found");
				return;
			}
			bkt.Connect = false;
        }
	}
}