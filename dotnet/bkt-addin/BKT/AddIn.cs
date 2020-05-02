using System.Diagnostics;
using Extensibility;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;

// xml validation
using System.Resources; // ResourceManager
using System.Xml;
using System.Xml.Schema;
using System.Xml.Linq;

// office libraries
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
// using Microsoft.Office.Tools;
// using Microsoft.Office.Tools.Ribbon;

// python scripting hosting
using Microsoft.Scripting.Hosting;
using IronPython.Hosting;


using System.Windows.Forms; // for MessageBox
using System.Threading; // for Thread.Sleep


// to determine host application and hook application events
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Visio = Microsoft.Office.Interop.Visio; //Note: Visio Interop need to be embedded!
using Outlook = Microsoft.Office.Interop.Outlook;

// Mouse/Key-Hook: IKeyboardMouseEvents, Hook
using Gma.System.MouseKeyHook;

internal enum HostApplication {Unknown=0, Excel, PowerPoint, Word, Visio, Outlook}




namespace BKT
{
    
    [ComVisible(true)]
    [ProgId("BKT.AddIn")]
    [Guid("8EA4071E-7BD4-48DA-B96D-21AD02E1C238")]
#if OFFICE2010
    public class AddIn : Extensibility.IDTExtensibility2, IRibbonExtensibility 
#else
    public class AddIn : Extensibility.IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
#endif
    {
        private object app;
        private ScriptEngine ipy;
        private ScriptScope scope;
        private dynamic python_delegate;
        private bool broken;
        private bool debug;
        private bool log_show_msgbox;
        private bool async_startup;
        private readonly int instance_id;
        private static int instance_id_counter = 0;
        private static int finalize_counter = 0;
        private static int connected_counter = 0;
        private static int disconnected_counter = 0;
        private Context context;
        private bool created;
        private string async_startup_ribbon_id;
        private IRibbonUI async_startup_ribbon;
        private TextWriterTraceListener listener;
        private FileStream logFileStream;
        
        private HostApplication host = HostApplication.Unknown;
        private string hostAppName;
        private IKeyboardMouseEvents m_GlobalHook;
        private bool keymouse_hooks_activated = false;
        
        private DateTime ppt_last_selection_changed = DateTime.MinValue;
        private int ppt_last_selection_shape_id = 0;
        

        #region Contructor and reset
        // ============================
        // = Constructor / Destructor =
        // ============================
        
        public AddIn() {
            instance_id_counter += 1;
            instance_id = instance_id_counter;
            
            ReloadConfig();
            bool log_write_file = bool.Parse(GetConfigEntry("log_write_file", "false"));
            
            // configure Debug logging
            Debug.Listeners.Clear();
            if ( log_write_file ) {
                string codebase = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
                string path = Path.Combine(codebase, "..", "bkt-debug.log");
                try
                {
                    logFileStream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);

                    TextWriterTraceListener listener;
                    listener = new TextWriterTraceListener(logFileStream);
                    Debug.Listeners.Add(listener);
                }
                catch(Exception ex)
                {
                    Console.WriteLine("Error creating FileStream for trace file \"{0}\":" +
                        "\r\n{1}", path, ex.Message);
                }
            }
            Debug.AutoFlush = true;
            Debug.WriteLine("");
            Debug.WriteLine("================================================================================");
            DebugMessage("Addin started");
            
            // initialize Mouse/Key-Hooks
            try
            {
                bool use_keymouse_hooks = Boolean.Parse(GetConfigEntry("use_keymouse_hooks", "true"));
                DebugMessage("Subscribe to Key/Mouse Events: " + use_keymouse_hooks.ToString());
                if (use_keymouse_hooks)
                {
                    HookEvents();
                }
            }
            catch (Exception)
            {
                DebugMessage("Error subscribing to Mouse Events");
            }
            
        }
        
        
        
        ~AddIn() {
            finalize_counter += 1;
        }
        
        public void Dispose() {
            DebugMessage("Dispose");
            
            UnhookEvents();
            
            // finalize debug messages
            if (listener != null) {
                listener.Flush();
                listener.Close();
                listener.Dispose();
                listener = null;
                logFileStream.Close();
                logFileStream.Dispose();
            }
        }
        
        private void Reset() {
            DebugMessage("Reset");
            
            created = false;
            app = null;
            ipy = null;
            scope = null;
            python_delegate = null;
            context = null;
            broken = false;
            debug = false;
            log_show_msgbox = false;
            async_startup_ribbon = null;
            
        }
        #endregion


        #region Logging
        // ===========
        // = Logging =
        // ===========
        
        private void Message(string s) {
            MessageBox.Show(s);
        }
        
        private void DebugMessage(string s) {
            Debug.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + ": " + s);
            // Debug.Flush(); --> not required as AutoFlush=true
        }

        private void LogMessage(string s) {
            DebugMessage(s);
            if(log_show_msgbox) {
                MessageBox.Show(s);
            }
        }
        #endregion
        
        
        #region Config
        // ==========
        // = Config =
        // ==========
        
        private Config config;
        
        public void ReloadConfig() {
            DebugMessage("ReloadConfig");
            string codebase = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
            string path = Path.Combine(codebase, "..\\config.txt");
            config = ConfigParser.Parse(path);
        }
        
        public string GetConfigEntry(string key) {
            return GetConfigEntry(key, null);
        }

        public string GetConfigEntry(string key, string default_value) {
            if(config.items.ContainsKey(key)) {
                return config.items[key];
            } else {
                return default_value;
            }
        }
        
        public string DumpConfig() {
            string s = "";
            foreach (string key in config.items.Keys) {
                s = s + key + " = " + GetConfigEntry(key) + "\r\n";
            }
            return s;
        }

        public string GetBuildConfiguration() {
#if OFFICE2010
            return "OFFICE2010";
#elif DEBUG
            return "DEBUG";
#else
            return "RELEASE";
#endif
        }

        public string GetBuildRevision() {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }
        #endregion
        
        
        #region Python engine
        // ===================
        // = Python delegate =
        // ===================
        
        private void LoadPython() {
            DebugMessage("LoadPython Called");
            
            
            // if(!IronPythonLoader.GetInstance().CheckIronPythonLocation()) {
            //     broken = true;
            //     Message("IronPython not found :(");
            // } else {
            //     LogMessage("IronPython located at " + IronPythonLoader.GetInstance().GetIronPythonLocation());
            try {
                LogMessage("Before CreatePythonEngine...");
                CreatePythonEngine();
                if(debug) {
                    TryLoadDebugger();
                }
            } catch (Exception e) {
                broken = true;
                Message(e.ToString());
            }
            // }
        }
        
        private void TryLoadDebugger() {
            try {
                Python.ImportModule(scope, "pydevd");
                //ipy.Execute("import pydevd", scope);
                ipy.Execute("pydevd.settrace(stdoutToServer=True, stderrToServer=True)", scope);
            } catch(Exception e) {
                debug = false;
                Message(e.ToString());
            }
        }
        
        public dynamic GetDelegate() {
            return python_delegate;
        }
                
        private void BootstrapAddIn() {
            DebugMessage("BootstrapAddIn called");
            Debug.Indent();
            
            // load config
            string ipy_addin_module = GetConfigEntry("ipy_addin_module");
            if(ipy_addin_module == null) {
                broken = true;
                Debug.Unindent();
                return;
            }
            
            // import bootstrap module
            DebugMessage("import module");
            // Python.ImportModule(ipy, ipy_addin_module);
            ipy.ImportModule(ipy_addin_module);
            dynamic module = ipy.GetSysModule().GetVariable("modules").get(ipy_addin_module);
            DebugMessage("done. loaded bootstrap module: " + ipy_addin_module);
            
            // start addin on python side
            try {
                DebugMessage("create addin on python side");
                python_delegate = module.create_addin();
                DebugMessage("done.");
            } catch (Exception e) {
                Message(e.ToString());
            }

            if(python_delegate == null) {
                Debug.Unindent();
                throw new Exception("addin bootstrapper returned null");
            } else if(!created) {
                // FIXME: check for context==null instead
                DebugMessage("create context object");
                context = new Context(app, this, this.config, debug, hostAppName);
                DebugMessage("calling on_create");
                python_delegate.on_create(context);
                DebugMessage("done.");
            }
            Debug.Unindent();
        }
        
        private void AsyncStartup() {
            DebugMessage("AsyncStartup called");
            try {
                if (!created && async_startup) {
                    BootstrapAddIn();
                    // PythonOnRibbonLoad aufrufen ?
                    // wird bei async starup vorher geblockt
                    // deswegen hier noch ribbon setzen
                    if (async_startup_ribbon != null) {
                        context.ribbon = async_startup_ribbon;
                    }
                    // customUI fuer naechsten Start neu laden
                    // stellt sicher, dass auf Python Seite alles aufgerufen wurde
                    GetPythonCustomUIAndWriteToFile(async_startup_ribbon_id);
                    //FIXME: check whether this invalidates at the right time
                    created = true;
                    if (context.ribbon != null) {
                        context.ribbon.Invalidate();
                    }
                }
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        private void CreatePythonEngine() {
            DebugMessage("CreatePythonEngine called");
            Debug.Indent();
            DebugMessage("get instance");
            
            // check python debugging
            bool debug_enabled = bool.Parse(GetConfigEntry("pydev_debug","false"));
            var options = new Dictionary<string, object>();
            if(debug_enabled) {
                options["Frames"] = true;
                options["FullFrames"] = true;
            }
            
            // initialze scripting engine
            // IronPython will load from: <root>\bin\
            DebugMessage("create scripting engine");
            ipy = Python.CreateEngine(options);
            DebugMessage("create scope");
            scope = ipy.CreateScope();
            
            // Initialize system path for Python
            // Python will load from:
            //   <root>\bin
            //   <root>\bin\Lib
            DebugMessage("add ironpython paths");
            //dynamic paths = Python.GetSysModule(ipy).GetVariable("path");
            ICollection<string> paths = ipy.GetSearchPaths();
            paths.Clear();
            string root = GetConfigEntry("ironpython_root");
            paths.Add(root); //path is required for loading wpf fluent ribbon resources
            paths.Add(Path.Combine(root,"Lib")); //path is standard python library
            //path.Add(Path.Combine(root,"DLLs"));
            //path.Add(Path.Combine(root,"site-packages"));
            
            string ipy_addin_path = GetConfigEntry("ipy_addin_path");
            if(ipy_addin_path != null) {
                DebugMessage("add ipy addin path");
                paths.Add(ipy_addin_path);
                LogMessage("addin_path " + ipy_addin_path + " added to sys.path");
            }
            DebugMessage("Python SysModule path= " + ipy_addin_path);
            
            // add debug path
            if(debug_enabled) {
                DebugMessage("add pydev-codebase (debug)");
                string pydev_codebase = GetConfigEntry("pydev_codebase");
                if(pydev_codebase == null) {
                    Message("debugging enabled, but pydev_codebase not set");
                } else {
                    debug = true;
                    paths.Add(pydev_codebase);
                }
            }

            ipy.SetSearchPaths(paths);

            //this.ipy = ipy;
            //this.scope = scope;
            Debug.Unindent();
            DebugMessage("CreatePythonEngine done.");
        }
        #endregion
        
        
        #region Addin interface
        // ==========================
        // = Shared Addin Interface =
        // ==========================
        
        public void OnConnection(object application, ext_ConnectMode connect_mode, object addin_inst, ref Array custom)  {
            connected_counter += 1;
            OnConnection2(application);
        }
        
        public void OnConnection2(object application)
        {
            DebugMessage("OnConnection2 called");
            
            try {
                // determine host and bind events
                DetermineHostApplication(application);
                
                // reset
                DebugMessage("ReloadConfig");
                ReloadConfig();
                this.Reset();
                this.app = application;
                // dump config
                log_show_msgbox = bool.Parse(GetConfigEntry("log_show_msgbox", "false"));
                string msg = "instance_id=" + instance_id + "\r\nfinalize_count=" + finalize_counter;
                msg += "\r\n\r\n" + DumpConfig();
                LogMessage(msg);
                // initialize python instance
                DebugMessage("Initialize Python instance");
                LoadPython();
                
                // FIXME: optional, je nach Einstellung, Bootstrap erst beim Aufruf der ersten Click-Aktion
                async_startup = bool.Parse(GetConfigEntry("async_startup", "false"));
                LogMessage("async_startup =" + async_startup );
                
                if (! async_startup) {
                    // bootstrap directly
                    BootstrapAddIn();
                    created = true;
                    
                } else {
                    LogMessage("Starting asynchronous Bootstrap");
                    // bootstrap asynchronously
                    Thread bootstrapperThread = new Thread(AsyncStartup);
                    bootstrapperThread.Start();
                }
                
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        public void DetermineHostApplication(object application)
        {
            try {
                if (application is Excel.Application)
                {
                    host = HostApplication.Excel;
                    hostAppName = ((Excel.Application)application).Name;
                    DebugMessage("host application: " + hostAppName);
                    BindExcelEvents((Excel.Application)application);
                }
                else if (application is PowerPoint.Application)
                {
                    host = HostApplication.PowerPoint;
                    hostAppName = ((PowerPoint.Application)application).Name;
                    DebugMessage("host application: " + hostAppName);
                    BindPowerPointEvents((PowerPoint.Application)application);
                }
                else if (application is Word.Application)
                {
                    host = HostApplication.Word;
                    hostAppName = ((Word.Application)application).Name;
                    DebugMessage("host application: " + hostAppName);
                    BindWordEvents((Word.Application)application);
                }
                else if (application is Visio.Application)
                {
                    host = HostApplication.Visio;
                    hostAppName = ((Visio.Application)application).Name;
                    DebugMessage("host application: " + hostAppName);
                    // BindVisioEvents((Visio.Application)application);
                }
                else if (application is Outlook.Application)
                {
                    host = HostApplication.Outlook;
                    hostAppName = ((Outlook.Application)application).Name;
                    DebugMessage("host application: " + hostAppName);
                    // BindOutlookEvents((Outlook.Application)application);
                }
                else
                {
                    DebugMessage("host application unknown");
                }
            } catch (Exception) {
                DebugMessage("error dertermining host application (maybe visio interop not installed)");
            }
        }

        private void UnbindHostApplicationEvents()
        {
            //proper unbinding of events to avoid error "Microsoft.CSharp.RuntimeBinder.RuntimeBinderException" after addin reload
            try {
                if (host == HostApplication.PowerPoint)
                {
                    UnbindPowerPointEvents((PowerPoint.Application)context.app);
                }
                else if (host == HostApplication.Excel)
                {
                    UnbindExcelEvents((Excel.Application)context.app);
                }
                else if (host == HostApplication.Word)
                {
                    UnbindWordEvents((Word.Application)context.app);
                }
                else
                {
                    DebugMessage("no unbinding for host application");
                }
            } catch (Exception) {
                DebugMessage("error unbinding host application events");
            }
        }
        
        
        public void OnDisconnection(ext_DisconnectMode remove_mode, ref Array custom)
        {    
            LogMessage("OnDisconnection: instance_id=" + instance_id);
            try {
                UnbindHostApplicationEvents();
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
            try {
                if(!broken) {
                    python_delegate.on_destroy();
                }
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
            try {
                if(ipy != null) {
                    ipy.Runtime.Shutdown();
                }
                Reset();
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
            disconnected_counter += 1;
            if (connected_counter == disconnected_counter) {
                Dispose();
            }
        }
        
        public void OnAddInsUpdate(ref Array custom)
        {
        }
        
        public void OnStartupComplete(ref Array custom)
        {
        }
        
        public void OnBeginShutdown(ref Array custom)
        {
            DebugMessage("OnBeginShutdown");
        }
        
        private string GetFilenameFromRibbonId(string ribbon_id) {
            // IronPythonLoader loader = IronPythonLoader.GetInstance();
            string dirname = GetConfigEntry("ipy_addin_path") + "\\resources\\xml\\";
            Directory.CreateDirectory(dirname);
            return dirname + ribbon_id + ".xml";            
        }
        
        
        private string GetPythonCustomUIAndWriteToFile(string ribbon_id) {
            if(broken) {
                return "";
            }
            try {
                string customUI = python_delegate.get_custom_ui(ribbon_id);
                //DebugMessage(customUI);
                string filename = GetFilenameFromRibbonId(ribbon_id);
                System.IO.StreamWriter file = new System.IO.StreamWriter(filename);
                file.WriteLine(customUI);
                file.Close();
                LogMessage("wrote xml in: " + filename);
#if DEBUG
                VerifyCustomUI(customUI);
#endif
                return customUI;
            } catch (Exception e) {
                broken = true;
                Message(e.ToString());
                return "";
            }
        }
        
        private string GetCustomUIFromFile(string ribbon_id) {
            string filename = GetFilenameFromRibbonId(ribbon_id);
            LogMessage("loading xml from: " + filename);
            string customUI;
            if (File.Exists(filename)) {
                StreamReader streamReader = new StreamReader(filename);
                customUI = streamReader.ReadToEnd();
                streamReader.Close();
                //LogMessage(customUI);
            } else {
                LogMessage("file not found.");
                customUI = "";
            }
            
            return customUI;
        }
        
        public string GetCustomUI(string ribbon_id) {
            DebugMessage("GetCustomUI called");
            
            // remeber ribbon_id for asynchronous startup
            async_startup_ribbon_id = ribbon_id;
            
            // if not in aync mode, load custom ui
            if(!async_startup) {
                return GetPythonCustomUIAndWriteToFile(ribbon_id);
            } 
            
            // from here on:
            // addin is in async mode
            
            if (! created) {
                // python scope not created yet, return custom ui from file
                return GetCustomUIFromFile(ribbon_id);
            }
            
            if(broken) {
                // fallback for broken addin
                return "";
            }
            
            // we're in async mode, addin is created, not broken
            try {
                // load custom ui
                return GetPythonCustomUIAndWriteToFile(ribbon_id);
            } catch (Exception e) {
                broken = true;
                Message(e.ToString());
                return "";
            }
        }
        
        
        public void VerifyCustomUI(string text)
        {
            ResourceManager resources = new ResourceManager("BKT.Properties.xml_schemata", Assembly.GetExecutingAssembly());
            byte[] xsd_data = (byte[]) resources.GetObject("customui14_xsd");
            Stream xsd_res = new MemoryStream(xsd_data);
            
            XmlSchemaSet ss = new XmlSchemaSet();
            ss.Add("http://schemas.microsoft.com/office/2009/07/customui", XmlReader.Create(xsd_res));
                // File.OpenRead(@"C:\Office 2010 Developer Resources\Schemas\customui14.xsd"))
            
            xsd_res.Close();
            
            XDocument doc = XDocument.Parse(text);
            DebugMessage("Validating XML...");
            doc.Validate(ss, new ValidationEventHandler(ValidationCallBack));
            DebugMessage("Validating XML completed!");
        }

        private void ValidationCallBack(object sender, ValidationEventArgs vea)
        {    
            DebugMessage("\t event sender: " + sender);
            DebugMessage("\t validation args: " + vea.Exception);
            DebugMessage("");
        }
        #endregion
        
        
        #region Mouse keyboard events
        // ========================
        // = Mouse and Key events =
        // ========================
        
        // Mouse and Key events are needed to hide context-dialogs, which are displayed top-most
        // - generally, if another application gets focus
        // - if mouse is clicked outside the context-dialog
        // - if mouse is clicked outside the office application (putting focus on another application)
        // - on keystrokes, enspecially alt-tab
        // Due to a bug in Office 2013 the key-events will not fire if office has focus.
        // On alt-tab only the key-up event will fire
        
        private void MouseDownEvent(object sender, MouseEventExtArgs e)
        {
            DebugMessage(String.Format("MouseDown: \t{0}; \t System Timestamp: \t{1}", e.Button, e.Timestamp));
            if (!created) return;
            ppt_last_selection_changed = DateTime.MinValue; // Reset timestamp for selection_changed event
            python_delegate.mouse_down(sender, e);
        }
        
        private void MouseUpEvent(object sender, MouseEventExtArgs e)
        {
            DebugMessage(String.Format("MouseUp: \t{0}; \t System Timestamp: \t{1}", e.Button, e.Timestamp));
            if (!created) return;
            python_delegate.mouse_up(sender, e);
        }

        // private void MouseMoveEvent(object sender, MouseEventExtArgs e)
        // {
        //     // DebugMessage(String.Format("MouseMove: \t{0} / {1}; \t System Timestamp: \t{2}", e.X, e.Y, e.Timestamp));
        //     if (!created) return;
        //     python_delegate.mouse_move(sender, e);
        // }

        private void MouseDoubleClickEvent(object sender, MouseEventArgs e)
        {
            DebugMessage(String.Format("MouseDoubleClick: \t{0}", e.Button));
            if (!created) return;
            python_delegate.mouse_double_click(sender, e);
        }

        private void MouseDragStartedEvent(object sender, MouseEventExtArgs e)
        {
            DebugMessage(String.Format("MouseDragStartedEvent: \t{0} / {1}; \t System Timestamp: \t{2}", e.X, e.Y, e.Timestamp));
            if (!created) return;
            python_delegate.mouse_drag_start(sender, e);
        }

        private void MouseDragFinishedEvent(object sender, MouseEventExtArgs e)
        {
            DebugMessage(String.Format("MouseDragFinishedEvent: \t{0} / {1}; \t System Timestamp: \t{2}", e.X, e.Y, e.Timestamp));
            if (!created) return;
            python_delegate.mouse_drag_end(sender, e);
        }

        private void KeyDownEvent(object sender, KeyEventArgs e)
        {
            DebugMessage(String.Format("KeyDown: \t{0}", e.KeyCode));
            if (!created) return;
            python_delegate.key_down(sender, e);
        }
                
        private void KeyUpEvent(object sender, KeyEventArgs e)
        {
            DebugMessage(String.Format("KeyUp: \t{0}", e.KeyCode));
            if (!created) return;
            python_delegate.key_up(sender, e);
        }
        
        public void HookEvents()
        {
            if (keymouse_hooks_activated)
                return;

            // m_GlobalHook = Hook.GlobalEvents(); //Global events -> leads to performance issues when starting ppt
            m_GlobalHook = Hook.AppEvents(); //Only application events
            m_GlobalHook.MouseDownExt += MouseDownEvent;
            m_GlobalHook.MouseUpExt += MouseUpEvent;
            m_GlobalHook.MouseDoubleClick += MouseDoubleClickEvent;
            // m_GlobalHook.MouseMoveExt += MouseMoveEvent;

            m_GlobalHook.MouseDragStartedExt += MouseDragStartedEvent;
            m_GlobalHook.MouseDragFinishedExt += MouseDragFinishedEvent;

            m_GlobalHook.KeyDown += KeyDownEvent;
            m_GlobalHook.KeyUp += KeyUpEvent;
            keymouse_hooks_activated = true;
        }

        public void UnhookEvents()
        {
            if (keymouse_hooks_activated)
            {
                m_GlobalHook.MouseDownExt -= MouseDownEvent;
                m_GlobalHook.MouseUpExt -= MouseUpEvent;
                m_GlobalHook.MouseDoubleClick -= MouseDoubleClickEvent;
                // m_GlobalHook.MouseMoveExt -= MouseMoveEvent;
                m_GlobalHook.MouseDragStartedExt -= MouseDragStartedEvent;
                m_GlobalHook.MouseDragFinishedExt -= MouseDragFinishedEvent;
                m_GlobalHook.KeyDown -= KeyDownEvent;
                m_GlobalHook.KeyUp -= KeyUpEvent;

                //It is recommened to dispose it
                m_GlobalHook.Dispose();
                m_GlobalHook = null;
            }
            keymouse_hooks_activated = false;
        }
        
        
        private void HookEventsCallback(IRibbonControl control)
        {
            HookEvents();
        }

        private void UnhookEventsCallback(IRibbonControl control)
        {
            UnhookEvents();
        }
        
        public bool GetMouseKeyHookActivated(IRibbonControl control)
        {
            return keymouse_hooks_activated;
        }
        
        public void ToggleMouseKeyHookActivation(IRibbonControl control, bool pressed)
        {
            if (keymouse_hooks_activated)
            {
                UnhookEvents();
            }
            else
            {
                HookEvents();
            }
        }
        #endregion
        
        
        #region Application events
        // ======================
        // = Application events =
        // ======================
        
        // EXCEL
        
        private Excel.AppEvents_WorkbookOpenEventHandler events_xls_open = null;
        private Excel.AppEvents_NewWorkbookEventHandler events_xls_new = null;
        private void BindExcelEvents(Excel.Application application)
        {
            events_xls_open = new Excel.AppEvents_WorkbookOpenEventHandler(Excel_WorkbookOpen);
            events_xls_new = new Excel.AppEvents_NewWorkbookEventHandler(Excel_NewWorkbook);
            ((Excel.AppEvents_Event)application).WorkbookOpen += events_xls_open;
            ((Excel.AppEvents_Event)application).NewWorkbook += events_xls_new;
        }

        private void UnbindExcelEvents(Excel.Application application)
        {
            ((Excel.AppEvents_Event)application).WorkbookOpen -= events_xls_open;
            ((Excel.AppEvents_Event)application).NewWorkbook -= events_xls_new;
            DebugMessage("Excel events unbinded");
        }
        
        private void Excel_NewWorkbook(Excel.Workbook workbook)
        {
            DebugMessage("Excel: new workbook: " + workbook.FullName);
            try {
#if OFFICE2010
#else
                if(workbook.Windows.Count > 0) {
                    CreateTaskPaneForWindow(workbook.Windows[1]);
                }
#endif
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
        }

        private void Excel_WorkbookOpen(Excel.Workbook workbook)
        {
            DebugMessage("Excel: workbook opened: " + workbook.FullName);
            try {
#if OFFICE2010
#else
                if(workbook.Windows.Count > 0) {
                    CreateTaskPaneForWindow(workbook.Windows[1]);
                }
#endif
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
        }
        
        // POWER POINT

        private PowerPoint.EApplication_PresentationOpenEventHandler events_ppt_presentation_open = null;
        private PowerPoint.EApplication_NewPresentationEventHandler events_ppt_new_presentation = null;
        private PowerPoint.EApplication_WindowSelectionChangeEventHandler events_ppt_win_sel_change = null;
        
        private void BindPowerPointEvents(PowerPoint.Application application)
        {
            events_ppt_presentation_open = new PowerPoint.EApplication_PresentationOpenEventHandler(PowerPoint_PresentatonOpen);
            events_ppt_new_presentation = new PowerPoint.EApplication_NewPresentationEventHandler(PowerPoint_NewPresentation);
            events_ppt_win_sel_change = new PowerPoint.EApplication_WindowSelectionChangeEventHandler(PowerPoint_WindowSelectionChange);
            ((PowerPoint.EApplication_Event)application).PresentationOpen += events_ppt_presentation_open;
            ((PowerPoint.EApplication_Event)application).NewPresentation += events_ppt_new_presentation;
            ((PowerPoint.EApplication_Event)application).WindowSelectionChange += events_ppt_win_sel_change;
        }

        private void UnbindPowerPointEvents(PowerPoint.Application application)
        {
            ((PowerPoint.EApplication_Event)application).PresentationOpen -= events_ppt_presentation_open;
            ((PowerPoint.EApplication_Event)application).NewPresentation -= events_ppt_new_presentation;
            ((PowerPoint.EApplication_Event)application).WindowSelectionChange -= events_ppt_win_sel_change;
            DebugMessage("Powerpoint events unbinded");
        }
        
        private void PowerPoint_NewPresentation(PowerPoint.Presentation presentation)
        {
            DebugMessage("PowerPoint: new presentation: " + presentation.FullName);
            try {
#if OFFICE2010
#else
                if(presentation.Windows.Count > 0) {
                    CreateTaskPaneForWindow(presentation.Windows[1]);
                }
#endif
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
        }

        private void PowerPoint_PresentatonOpen(PowerPoint.Presentation presentation)
        {
            DebugMessage("PowerPoint: presentation opened: " + presentation.FullName);
            try {
#if OFFICE2010
#else
                if(presentation.Windows.Count > 0) {
                    CreateTaskPaneForWindow(presentation.Windows[1]);
                }
#endif
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
        }

        private void PowerPoint_WindowSelectionChange(PowerPoint.Selection selection)
        {
            DebugMessage("PowerPoint: window selection changed instance="+instance_id);
            try {
                selection_type = (int)selection.Type;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText) {
                    int shape_id = 0;
                    
                    // Set values for fast enable events
                    selection_shapes = 1;
                    selection_containstextframe = true;

                    // Store shape id
                    if (selection.HasChildShapeRange) {
                        shape_id = selection.ChildShapeRange[1].Id;
                    } else {
                        shape_id = selection.ShapeRange[1].Id;
                    }

                    // Selection is changed for each key press (e.g. while typing), therefore we introduce some thresholds:
                    // Update timestamp after 2 seconds, when shape id changed, or (see MouseDownEvent) when mouse is clicked
                    if ((DateTime.Now-ppt_last_selection_changed).TotalSeconds > 2 || ppt_last_selection_shape_id != shape_id) {
                        ppt_last_selection_changed  = DateTime.Now;
                        ppt_last_selection_shape_id = shape_id;
                    } else {
                        // stop method here, no python_delegate (at the bottom)
                        return;
                    }
                } else {
                    ppt_last_selection_changed = DateTime.MinValue;
                    // Set values for fast enable events
                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) {
                        if (selection.HasChildShapeRange) {
                            selection_shapes = selection.ChildShapeRange.Count;
                        } else {
                            selection_shapes = selection.ShapeRange.Count;
                        }
                        selection_containstextframe = Ppt_Selection_Contains_Textframe(selection);
                    } else {
                        selection_shapes = 0;
                        selection_containstextframe = false;
                    }
                }

                python_delegate.ppt_selection_changed(selection);
            } catch (Exception e) {
                // On error (e.g. unsupported ViewType or ActivePane other than normal), set all value to nothing selected
                selection_type = 0; //PowerPoint.PpSelectionType.ppSelectionNone
                selection_shapes = 0;
                selection_containstextframe = false;
                LogMessage(e.ToString());
            }
        }

        private bool Ppt_Selection_Contains_Textframe(PowerPoint.Selection selection)
        {
            try {
                if (selection.HasChildShapeRange) {
                    if (selection.ChildShapeRange.HasTextFrame == MsoTriState.msoFalse) {
                        return false;
                    } else {
                        return true;
                    }
                } else {
                    if (selection.ShapeRange.HasTextFrame != MsoTriState.msoFalse || selection.ShapeRange.HasTable != MsoTriState.msoFalse) {
                        return true;
                    } else {
                        foreach (PowerPoint.Shape el in selection.ShapeRange) {
                            if (el.Type == MsoShapeType.msoGroup && el.GroupItems.Range(null).HasTextFrame != MsoTriState.msoFalse) {
                                return true;
                            }
                            if (el.Type == MsoShapeType.msoSmartArt) {
                                return true;
                            }
                        }
                        return false;
                    }
                }
            } catch (Exception) {
                // LogMessage(e.ToString());
                return false;
            }
        }
        
        // WORD
        
        private Word.ApplicationEvents4_DocumentOpenEventHandler events_word_open = null;
        private Word.ApplicationEvents4_NewDocumentEventHandler events_word_new = null;
        private void BindWordEvents(Word.Application application)
        {
            events_word_open = new Word.ApplicationEvents4_DocumentOpenEventHandler(Word_DocumentOpen);
            events_word_new = new Word.ApplicationEvents4_NewDocumentEventHandler(Word_NewDocument);
            ((Word.ApplicationEvents4_Event)application).DocumentOpen += events_word_open;
            ((Word.ApplicationEvents4_Event)application).NewDocument += events_word_new;
        }
        private void UnbindWordEvents(Word.Application application)
        {
            ((Word.ApplicationEvents4_Event)application).DocumentOpen -= events_word_open;
            ((Word.ApplicationEvents4_Event)application).NewDocument -= events_word_new;
            DebugMessage("Word events unbinded");
        }
        
        private void Word_NewDocument(Word.Document document)
        {
            DebugMessage("Word: new document: " + document.FullName);
            try {
#if OFFICE2010
#else
                if(document.Windows.Count > 0) {
                    CreateTaskPaneForWindow(document.Windows[1]);
                }
#endif
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
        }

        private void Word_DocumentOpen(Word.Document document)
        {
            DebugMessage("Word: document opened: " + document.FullName);
            try {
#if OFFICE2010
#else
                if(document.Windows.Count > 0) {
                    CreateTaskPaneForWindow(document.Windows[1]);
                }
#endif
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
        }
        #endregion


        #region Window handles
        // ===============================
        // = Get window handle for forms =
        // ===============================

        public object GetActiveWindow()
        {
            
            if (host == HostApplication.Excel)
            {
                return ((Excel.Application)context.app).ActiveWindow;
            }
            else if (host == HostApplication.PowerPoint)
            {
                if ( ((PowerPoint.Application)context.app).Windows.Count == 0) {
                    // Avoid error: System.Runtime.InteropServices.COMException (0x80048240): Application (unknown member) : Invalid request.  There is no currently active document window.
                    DebugMessage("GetActiveWindow: no active Windows!");
                    throw new NullReferenceException("No active windows");
                }
                return ((PowerPoint.Application)context.app).ActiveWindow;
            }
            else if (host == HostApplication.Word)
            {
                return ((Word.Application)context.app).ActiveWindow;
            }
            // else if (host == HostApplication.Visio)
            // {
            //     return ((Visio.Application)context.app).ActiveWindow;
            // }
            else
            {
                throw new NotSupportedException("Unknown host application");
            }
        }


        public int GetWindowHandle(object window=null)
        {
#if OFFICE2010
            return 0;
#else
            try
            {
                if (window == null)
                {
                    window = GetActiveWindow();
                }
                if (host == HostApplication.Excel)
                {
                    int windowID = ((Excel.Window)window).Hwnd;
                    return windowID;
                }
                else if (host == HostApplication.PowerPoint)
                {
                    int windowID = ((PowerPoint.DocumentWindow)window).HWND;
                    return windowID;
                }
                else if (host == HostApplication.Word)
                {
                    int windowID = ((Word.Window)window).Hwnd;
                    return windowID;
                }
                // else if (host == HostApplication.Visio)
                // {
                //     Int32 windowID = ((Visio.Window)window).WindowHandle32;
                //     return windowID;
                // }
                else
                {
                    return 0;
                }
            }
            catch (Exception)
            {
                return 0;
            }
#endif
        }

        public IWin32Window GetW32WindowHandle(object window=null)
        {
#if OFFICE2010
            return null;
#else 
            return new Win32Window(GetWindowHandle(window));
#endif
        }
        #endregion


        #region Taskpane stuff
        // =======================
        // = Task Pane Interface =
        // =======================

#if OFFICE2010
#else
        private ICTPFactory myCtpFactory;
        private CustomTaskPane myPane;
        private TaskPane myControl;
        
        private Dictionary<int, CustomTaskPane> TaskPanes = new Dictionary<int, CustomTaskPane>();
        

        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            DebugMessage("CTPFactoryAvailable called");
            // TODO: Initialisierung der Config aus dem IronPythonLoader in AddIn verlagern
            if (bool.Parse(GetConfigEntry("task_panes", "false")) == false) {
                return;
            }

            myCtpFactory = CTPFactoryInst;
            
            DebugMessage("Load fluent ribbon assembly");
            string codebase = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
            string path = Path.Combine(codebase, "Fluent.dll");
            Assembly.LoadFrom(path);
            
        }
        
        private CustomTaskPane CreateTaskPaneForWindow(object window)
        {
            if (myCtpFactory == null) {
                return null;
            }
            
            // TODO: task-pane xml aus python Welt laden
            // TODO: task-pane label aus xml lesen
            DebugMessage("Obtain task pane for window");
            try {
                
                DebugMessage("Remove orphaned Task Panes");
                RemoveOrphanedTaskPanes();
                
                DebugMessage("Check existing task panes");
                int windowID = GetWindowHandle(window);
                if (windowID != 0)
                {
                    if (TaskPanes.ContainsKey(windowID) == true) {
                        DebugMessage("Task pane exists. done");
                        return TaskPanes[windowID];
                    }
                }
                else if (myPane != null)
                {
                    DebugMessage("Task pane exists. done");
                    return myPane;
                }

                
                DebugMessage("Create new task panes");
                // create task pane and the custom-control within
                if (window == null)
                    myPane = myCtpFactory.CreateCTP("BKT.TaskPane", "BKT Task Pane", Type.Missing);
                else
                    myPane = myCtpFactory.CreateCTP("BKT.TaskPane", "BKT Task Pane", window);
                if (windowID != 0)
                {
                    DebugMessage("Remeber Task pane for window: " + windowID);
                    TaskPanes.Add(windowID, myPane);
                }
                
                myPane.VisibleStateChange += new _CustomTaskPaneEvents_VisibleStateChangeEventHandler(taskPane_VisibleChanged);
                // TODO: event for dock position changed
                
                myPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                myControl = (TaskPane)myPane.ContentControl;
                
                // give task pane control to python-delegate to route events
                myControl.WpfControl.SetPythonDelegate(python_delegate);
                myControl.WpfControl.UpdateContent();
                
                // TODO: visibility der task-pane in config speichern/ laden
                myPane.Visible = false;
                
                
                DebugMessage("Done with task pane");
                return myPane;
                
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
            return null;
        }
        
        private void RemoveOrphanedTaskPanes()
        {
            List<int> orphanedKeys = new List<int>();
            
            try {
                // iterate through Task Panes
                foreach (KeyValuePair<int, CustomTaskPane> pair in TaskPanes)
                {
                    CustomTaskPane ctp = pair.Value;
                    
                    // obtain Task Pane Window
                    object window;
                    try {
                        window = ctp.Window;
                    } catch (Exception) {
                        // System.Runtime.InteropServices.COMException (0x800A01A8): Der Aufgabenbereich wurde gelöscht oder ist anderweitig nicht mehr gültig.
                        window = null;
                    }
                    
                    // if no window was found, remember key
                    if (window == null)
                    {
                        ctp = null;
                        orphanedKeys.Add(pair.Key);
                    }
                }
                
                // remove Task Panes
                foreach (int key in orphanedKeys) 
                    TaskPanes.Remove(key);
                
                
                
            } catch (Exception e) {
                LogMessage(e.ToString());
            }
        }
        
        public bool GetPressed_TaskPaneToggler(IRibbonControl control)
        {
            if (bool.Parse(GetConfigEntry("task_panes", "false")) == false)
                return false;
            
            try
            {
                CustomTaskPane ctp = GetActiveTaskPane();
                if (ctp !=null)
                    return ctp.Visible;
                else
                    return false;
                
            } catch (Exception e) {
                DebugMessage(e.ToString());
                return false;
            }
        }

        public void OnAction_TaskPaneToggler(IRibbonControl control, bool pressed)
        {
            if (bool.Parse(GetConfigEntry("task_panes", "false")) == false) {
                return ;
            }
            try {
                CustomTaskPane ctp = GetActiveTaskPane();
                if (ctp != null)
                    ctp.Visible = pressed;
            } catch (Exception e) {
                DebugMessage(e.ToString());
            }
        }
        
        
        public CustomTaskPane GetActiveTaskPane()
        {
            try
            {
                object activeWindow = GetActiveWindow();
                int windowID = GetWindowHandle(activeWindow);
                if (TaskPanes.ContainsKey(windowID) == true)
                {
                    return TaskPanes[windowID];
                }
                else
                {
                    DebugMessage(string.Format("WARNING: Didn't find task pane for active window, hwnd={0}! Creating new task pane.", windowID));
                    return CreateTaskPaneForWindow(activeWindow);
                }
            }
            catch (NullReferenceException)
            {
                //no active window error
                return null;
            }
            catch (NotSupportedException)
            {
                //unknown app error
                if (myPane == null)
                {
                    return CreateTaskPaneForWindow(null);
                }
                else
                {
                    return myPane;
                }
            }
        }
        
        private void taskPane_VisibleChanged(CustomTaskPane CustomTaskPaneInst)
        {
            if (context.ribbon != null) {
                context.ribbon.Invalidate();
                //context.ribbon.InvalidateControl(xxx);
            }
        }
        
#endif
        #endregion
        
        
        #region Callbacks with Python delegation
        // ========================================
        // = Python Events: information callbacks =
        // ========================================
        // For callback signatures see
        //  https://msdn.microsoft.com/en-us/library/aa722523(v=office.12).aspx
        //  https://msdn.microsoft.com/en-us/library/bb736142(v=office.12).aspx
        
        public string PythonGetContent(IRibbonControl control)
        {    
            DebugMessage("event GetContent " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_content(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public string PythonGetDescription(IRibbonControl control)
        {    
            DebugMessage("event GetDescription " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_description(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public bool PythonGetEnabled(IRibbonControl control)
        {
            DebugMessage("event GetEnabled " + control.Id);
            if (!created) return false;
            try {
                return (python_delegate.get_enabled(control) == true);
            } catch (Exception e) {
                Message(e.ToString());
                return false;
            }
        }
         
        public Bitmap PythonGetImage(IRibbonControl control) {
            DebugMessage("event GetImage " + control.Id);
            if (!created) return null;
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_image(control);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public string PythonGetKeytip(IRibbonControl control)
        {    
            DebugMessage("event GetKeytip " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_keytip(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public string PythonGetLabel(IRibbonControl control)
        {    
            DebugMessage("event GetLabel " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_label(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public bool PythonGetPressed(IRibbonControl control)
        {
            DebugMessage("event GetPressed " + control.Id);
            if (!created) return false;
            try {
                return (python_delegate.get_pressed(control) == true);
            } catch (Exception e) {
                Message(e.ToString());
                return false;
            }
        }
        
        public string PythonGetScreentip(IRibbonControl control)
        {    
            DebugMessage("event GetScreentip " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_screentip(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public bool PythonGetShowImage(IRibbonControl control)
        {
            DebugMessage("event GetShowImage " + control.Id);
            if (!created) return false;
            try {
                return (python_delegate.get_show_image(control) == true);
            } catch (Exception e) {
                Message(e.ToString());
                return false;
            }
        }
        
        public bool PythonGetShowLabel(IRibbonControl control)
        {
            DebugMessage("event GetShowLabel " + control.Id);
            if (!created) return false;
            try {
                return (python_delegate.get_show_label(control) == true);
            } catch (Exception e) {
                Message(e.ToString());
                return false;
            }
        }
        
        public int PythonGetSize(IRibbonControl control)
        {
            DebugMessage("event GetSize " + control.Id);
            if (!created) return 0;
            try {
                var result = python_delegate.get_size(control);
                if(result == "large") {
                    return 1;
                } else {
                    return 0;
                }
            } catch (Exception e) {
                Message(e.ToString());
                return 0;
            }
        }
        
        public string PythonGetSupertip(IRibbonControl control)
        {    
            DebugMessage("event GetSupertip " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_supertip(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public string PythonGetText(IRibbonControl control)
        {    
            DebugMessage("event GetText " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_text(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }

        public string PythonGetTitle(IRibbonControl control)
        {    
            DebugMessage("event GetTitle " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_title(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public bool PythonGetVisible(IRibbonControl control)
        {
            DebugMessage("event GetVisible " + control.Id);
            if (!created) return false;
            try {
                return (python_delegate.get_visible(control) == true);
            } catch (Exception e) {
                Message(e.ToString());
                return false;
            }
        }
        
        
        // ====================================
        // = Python Events: gallery/combo box =
        // ====================================
        
        public int PythonGetItemCount(IRibbonControl control) {
            DebugMessage("event GetItemCount " + control.Id);
            if (!created) return 0;
            if(broken) {
                return 0;
            }
            try {
                var v = python_delegate.get_item_count(control);
                if (v==null) {
                    return 0;
                } else {
                    return v;
                }
            } catch (Exception e) {
                Message(e.ToString());
                return 0;
            }
        }

        public int PythonGetSelectedItemIndex(IRibbonControl control) {
            DebugMessage("event GetSelectedItemIndex " + control.Id);
            if (!created) return 0;
            if(broken) {
                return 0;
            }
            try {
                return python_delegate.get_selected_item_index(control);
            } catch (Exception e) {
                Message(e.ToString());
                return 0;
            }
        }
        
        public string PythonGetSelectedItemID(IRibbonControl control)
        {    
            DebugMessage("event GetSelectedItemID " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_selected_item_id(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        

        // ==============================================
        // = Python Events: gallery/combo box (indexed) =
        // ==============================================
        
        public int PythonGetItemHeight(IRibbonControl control) {
            DebugMessage("event GetItemHeight " + control.Id);
            if (!created) return 0;
            if(broken) {
                return 0;
            }
            try {
                return python_delegate.get_item_height(control);
            } catch (Exception e) {
                Message(e.ToString());
                return 0;
            }
        }
        
        public string PythonGetItemID(IRibbonControl control, int index) {
            DebugMessage("event GetItemID " + control.Id);
            if (!created) return "";
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_item_id(control, index);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public Bitmap PythonGetItemImage(IRibbonControl control, int index) {
        //public stdole.IPictureDisp GetItemImage(IRibbonControl oRbnCtrl, int iItemIndex)
            DebugMessage("event GetItemImage " + control.Id);
            if (!created) return null;
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_item_image(control, index);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public string PythonGetItemLabel(IRibbonControl control, int index) {
            DebugMessage("event GetItemLabel " + control.Id);
            if (!created) return "";
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_item_label(control, index);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public string PythonGetItemScreentip(IRibbonControl control, int index) {
            DebugMessage("event GetItemScreentip " + control.Id);
            if (!created) return "";
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_item_screentip(control, index);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public string PythonGetItemSupertip(IRibbonControl control, int index) {
            DebugMessage("event GetItemSupertip " + control.Id);
            if (!created) return "";
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_item_supertip(control, index);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public int PythonGetItemWidth(IRibbonControl control) {
            DebugMessage("event GetItemWidth " + control.Id);
            if (!created) return 0;
            if(broken) {
                return 0;
            }
            try {
                return python_delegate.get_item_width(control);
            } catch (Exception e) {
                Message(e.ToString());
                return 0;
            }
        }
        
        
        // ===================================
        // = Python Events: action callbacks =
        // ===================================
        
        public void PythonOnAction(IRibbonControl control)
        {    
            DebugMessage("event OnAction " + control.Id);
            if (!created) return;
            try {
                python_delegate.on_action(control);
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        public void PythonOnActionRepurposed(IRibbonControl control, ref bool cancelDefault)
        {    
            DebugMessage("event OnActionRepurposed " + control.Id);
            if (!created) return;
            try {
                cancelDefault = Convert.ToBoolean(python_delegate.on_action_repurposed(control));
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        public void PythonOnActionIndexed(IRibbonControl control, string selectedItem, int index)
        {    
            DebugMessage("event OnActionIndex " + control.Id);
            if (!created) return;
            try {
                python_delegate.on_action_indexed(control, selectedItem, index);
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        public void PythonOnToggleAction(IRibbonControl control, bool pressed)
        {    
            DebugMessage("event OnToggleAction " + control.Id);
            if (!created) return;
            try {
                python_delegate.on_toggle_action(control, pressed);
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        public void PythonOnChange(IRibbonControl control, string value)
        {    
            DebugMessage("event OnChange " + control.Id);
            if (!created) return;
            try {
                python_delegate.on_change(control, value);
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        
        // ======================================
        // = Python Events: image / ribbon load =
        // ======================================
        
        public Bitmap PythonLoadImage(string image_name) {
            DebugMessage("event LoadImage " + image_name);
            if (!created) return null;
            if(broken) {
                return null;
            }
            try {
                return python_delegate.load_image(image_name);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public void PythonOnRibbonLoad(IRibbonUI ui)
        {
            DebugMessage("event OnRibbonLoad");
            if (!created) 
            {
                async_startup_ribbon = ui;
                return;
            }
            try {
                context.ribbon = ui;
                python_delegate.on_ribbon_load(ui);
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        #endregion

        
        #region Fast enabled events
        // =======================
        // = Fast Enabled-Events =
        // =======================
        
        private int selection_type = 0;
        private int selection_shapes = 0;
        private bool selection_containstextframe = false;

        public bool GetEnabled_True(IRibbonControl control)
        {
            DebugMessage("event GetEnabled_True " + control.Id);
            if (!created) return false;
            return true;
        }

        public bool GetEnabled_Ppt_ShapesOrText(IRibbonControl control)
        {
            DebugMessage("event GetEnabled_Ppt_ShapesOrText " + control.Id);
            if (!created) return false;
            return selection_type == 2 || selection_type == 3;
        }

        public bool GetEnabled_Ppt_Shapes_ExactOne(IRibbonControl control)
        {
            DebugMessage("event GetEnabled_Ppt_Shapes_ExactOne " + control.Id);
            if (!created) return false;
            return selection_shapes == 1;
        }

        public bool GetEnabled_Ppt_Shapes_ExactTwo(IRibbonControl control)
        {
            DebugMessage("event GetEnabled_Ppt_Shapes_ExactTwo " + control.Id);
            if (!created) return false;
            return selection_shapes == 2;
        }

        public bool GetEnabled_Ppt_Shapes_MinTwo(IRibbonControl control)
        {
            DebugMessage("event GetEnabled_Ppt_Shapes_MinTwo " + control.Id);
            if (!created) return false;
            return selection_shapes >= 2;
        }

        public bool GetEnabled_Ppt_ContainsTextFrame(IRibbonControl control)
        {
            DebugMessage("event GetEnabled_Ppt_ContainsTextFrame " + control.Id);
            if (!created) return false;
            return selection_containstextframe;
        }
        
        /*
        Enabled based on selection
            0 = ppSelectionNone
            1 = ppSelectionSlide
            2 = ppSelectionShape
            3 = ppSelectionText
        */
        // public Boolean GetEnabled_Shapes_Selected(IRibbonControl control)
        // {
        //             DebugMessage("event GetEnabled_Selection_Available " + control.Id);
        //             if (!created) return false;
        //             return ((PowerPoint.Application)app.ActiveWindow.selection.Type == 2);
        // }
        //
        // public Boolean GetEnabled_Text_Selected(IRibbonControl control)
        // {
        //             DebugMessage("event GetEnabled_Selection_Available " + control.Id);
        //             if (!created) return false;
        //             return ((PowerPoint.Application)app.ActiveWindow.selection.Type == 3);
        // }

        #endregion
        
    }
}
