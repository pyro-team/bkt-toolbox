
using Microsoft.Office.Core;

namespace BKT
{
    /// <summary>
    /// Description of Class1.
    /// </summary>
    public class Context
	{	
		public object app {get; private set; }
		public AddIn addin {get; private set; }
		public bool debug {get; private set; }
		public IRibbonUI ribbon {get; internal set; }
		public Config config {get; private set; }
		public string hostAppName {get; private set; }
		
		public Context(object app, AddIn addin, Config config, bool debug, string hostAppName)
		{
			this.app = app;
			this.addin = addin;
			this.debug = debug;
			this.config = config;
			this.hostAppName = hostAppName;
		}
	}
}
