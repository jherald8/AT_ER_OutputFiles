using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace AT_ER_OutputFiles
{
    internal class ExecuteScript : FileTool
    {
        string projectDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        public void RunSapScripting(string script)
        {
            Console.WriteLine("\nExecuting Reports");
            string sapScript = $@"{projectDirectory}\{script}.vbs";
            ExecutingVB(null, null, sapScript);
            Console.WriteLine("\nDone Executing Reports");
        }
        public void LoginSapGui()
        {
            string sapScript = $@"{projectDirectory}\SAPautomation.py";
            string system = SystemIdentify();
            Console.WriteLine("\nSAPGUI Login");
            ExecutingPython(system, null, sapScript);
            Console.WriteLine("\nLogon Completed!");
        }
        public void DownloadGmail()
        {
            string gmailScript = $@"{projectDirectory}\GmailDownload.py";
            string gmailUser = ConfigurationSettings.AppSettings["GmailUser"];
            string gmailPass = ConfigurationSettings.AppSettings["GmailPass"];
            Console.WriteLine("\nStarting Downloading of Attachment in GMAIL");
            ExecutingPython(gmailUser, gmailPass, gmailScript);
            Console.WriteLine("\nDownload Completed!");
        }
        public void ExecutingVB(string param1, string param2, string script)
        {
            Process process = new Process();
            try
            {
                process.StartInfo.FileName = "cscript.exe";
                process.StartInfo.Arguments = "//B //Nologo \"" + script + "\"";
                process.Start(); // Start the process

                // Wait for the process to exit
                process.WaitForExit();

                Console.WriteLine("\nVBScript execution completed.");
            }
            catch (Exception ex)
            { Console.WriteLine("Error: " + ex.Message); }
            finally 
            { process.Close(); }
        }
        public void ExecutingPython(string param1, string param2, string script)
        {
            string pType = ConfigurationSettings.AppSettings["ProcessType"];
            
            string pythonPath = @$"{projectDirectory}\Python311\python.exe";
            string temp = ConfigurationSettings.AppSettings["Temp"];

            string comparisonPath = ConfigurationSettings.AppSettings["SourcePath"];
            string encodingPath = ConfigurationSettings.AppSettings["EncodingPath"];

            if (pType == "1")
                temp = comparisonPath.Replace(@"\", @"/");
            else if (pType == "2")
                temp = encodingPath.Replace(@"\", @"/");
            else if (pType == "4")
                temp = temp.Replace(@"\", @"/");

            string[] argumentToSend = { param1, param2, temp };

            string concatenatedStrings = string.Join(",", argumentToSend);

            // Create a ProcessStartInfo object to configure the process
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = pythonPath,
                Arguments = $"\"{script}\" \"{concatenatedStrings}\"", // Pass the path to your Python script as an argument
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            // Create a new process
            Process process = new Process
            {
                StartInfo = startInfo
            };

            // Start the process
            process.Start();

            // Read the output and error streams
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();
            // Wait for the process to exit
            process.WaitForExit();
        }
    }
}
