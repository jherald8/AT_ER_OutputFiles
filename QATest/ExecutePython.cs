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
    internal class ExecutePython
    {
        public void ExecutingPython()
        {
            string pType = ConfigurationSettings.AppSettings["ProcessType"];
            string comparisonPath = ConfigurationSettings.AppSettings["SourcePath"];

            string executablePath = Assembly.GetEntryAssembly().Location;
            string projectDirectory = Path.GetDirectoryName(executablePath);
            string pythonPath = @$"{projectDirectory}\Python311\python.exe";
            string scriptPath = $@"{projectDirectory}\GmailDownload.py";
            string gmailUser = ConfigurationSettings.AppSettings["GmailUser"];
            string gmailPass = ConfigurationSettings.AppSettings["GmailPass"];
            string temp = ConfigurationSettings.AppSettings["Temp"];
            string encodingPath = ConfigurationSettings.AppSettings["EncodingPath"];

            if (pType == "1")
                temp = comparisonPath.Replace(@"\", @"/");
            else if (pType == "2")
                temp = encodingPath.Replace(@"\", @"/");
            else if (pType == "4")
                temp = temp.Replace(@"\", @"/");

            string[] argumentToSend = { gmailUser, gmailPass, temp };

            string concatenatedStrings = string.Join(",", argumentToSend);

            Console.WriteLine("\nStarting Downloading of Attachment in GMAIL");

            // Create a ProcessStartInfo object to configure the process
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = pythonPath,
                Arguments = $"\"{scriptPath}\" \"{concatenatedStrings}\"", // Pass the path to your Python script as an argument
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

            Console.WriteLine("\nDownloading Completed!");
        }
    }
}
