using System.IO;
using System.Text;
using System.Collections.Generic;

namespace BKT
{
    public class Config {
        public Dictionary<string,string> items;
        public List<string> pythonpath;
        public List<string> modules;
        
        public Config() {
            items = new Dictionary<string, string>();
            pythonpath = new List<string>();
            modules = new List<string>();
        }
    }
    /// <summary>
    /// Description of Class1.
    /// </summary>
    public class ConfigParser
    {
            
        private ConfigParser()
        {
        }
        
        public static Config Parse(string path) {
            StreamReader reader = null;
            try {
                reader = new StreamReader(path, UTF8Encoding.UTF8);
                var config = new Config();
                while(true) {
                    string line = reader.ReadLine();
                    if(line == null) {
                        break;
                    }
                    line = line.Trim();
                    if(line.Length == 0) {
                        continue;
                       }
                    if(line.StartsWith("#")) {
                        // ignore comments
                        continue;
                    }
                    if(line.StartsWith("[")) {
                        // ignore section headers
                        continue;
                    }
                    int index = line.IndexOf('=');
                    if(index < 0) {
                        // ignore multiline values
                        continue;
                        //throw new IOException("illegal line format");
                    }
                    
                    string key = line.Substring(0,index).Trim();
                    string value = line.Substring(index+1).Trim();
                    if(key == "pythonpath") {
                        config.pythonpath.Add(value);
                    } else if (key == "module") {
                        config.modules.Add(value);
                    } else {
                        config.items.Add(key, value);
                    }
                }
                return config;
            } finally {
                if(reader != null) {
                    reader.Close();
                }
            }
        }
    }
}
