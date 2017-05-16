using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CaptureReg
{
    class Program
    {
        public class ComRegistrations
        {
            public string DllDirectory { get; set; }
            public List<string> NativeCOMDLLs { get; set; }
            public List<string> ComAssemblies { get; set; }
        }

        [DllImport("RegistryCaptureHelper.dll")]
        static extern bool StartCapture();

        [DllImport("RegistryCaptureHelper.dll")]
        static extern bool StopCapture();

        [DllImport("RegistryCaptureHelper.dll")]
        static extern bool UnloadHive();

        [DllImport("RegistryCaptureHelper.dll", CharSet = CharSet.Auto)]
        static extern bool TestRegWrites(String regKey);

        [DllImport("RegistryCaptureHelper.dll", CharSet = CharSet.Auto)]
        static extern bool RegisterNativeCOMDll(String dllFileName);

        [DllImport("RegistryCaptureHelper.dll", CharSet = CharSet.Auto)]
        static extern bool PatchNativeCOMDllPaths(String DllDirectory);

        static int Main(string[] args)
        {
            if (args.Length >= 1 && args[0] == "-capture") {
                var comReg = ComRegistrationHelper.ReadConfiguration(args[1]);
                StartCapture();
                RegisterNativeCOMDlls(comReg);
                PatchNativeCOMDllPaths(comReg.DllDirectory);
                RegisterCOMAssemblies(comReg);
                return StopCapture() ? 0 : 1;
            }

            File.Copy("TemplateHive.dat", "ActiveHive.dat", true);
            var captureArgs = args.Length >= 1 ? args[0] : "comreg.json";
            var ret = RunProcess(Process.GetCurrentProcess().MainModule.FileName, "-capture " + captureArgs, true);
            if (ret == 1) {
                UnloadHive();
            }
            return 0;
        }

        private static void RegisterNativeCOMDlls(ComRegistration comReg)
        {
            foreach (var fileName in comReg.NativeCOMDLLs) {
                RegisterNativeCOMDll(comReg.DllDirectory + fileName);
            }
        }

        private static void RegisterCOMAssemblies(ComRegistration comReg)
        {
            RegistrationServices RegistrationServices = new RegistrationServices();

            foreach (var fileName in comReg.ComAssemblies) {
                RegisterCOMAssembly(RegistrationServices, comReg.DllDirectory + fileName);
            }
        }

        private static void RegisterCOMAssembly(RegistrationServices regSvcs, string assemblyFileName)
        {
            var asm = Assembly.UnsafeLoadFrom(assemblyFileName);
            regSvcs.RegisterAssembly(asm, AssemblyRegistrationFlags.None);
        }

        private static int RunProcess(string appName, string args, bool block = false)
        {
            int ret = 0;
            Process process = new Process();
            process.StartInfo.FileName = appName;
            process.StartInfo.Arguments = args;
            process.StartInfo.CreateNoWindow = true;
            process.Start();
            if (block) {
                process.WaitForExit();
                ret = process.ExitCode;
            }
            return ret;
        }
    }
}
