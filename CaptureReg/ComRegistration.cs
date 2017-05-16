using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaptureReg
{
    public class ComRegistration
    {
        public string DllDirectory { get; set; }
        public List<string> NativeCOMDLLs { get; set; }
        public List<string> ComAssemblies { get; set; }
    }

    public class ComRegistrationHelper
    {
        public static ComRegistration ReadConfiguration(string configFileName)
        {
            return JsonConvert.DeserializeObject<ComRegistration>(File.ReadAllText(configFileName));
        }
    }
}
