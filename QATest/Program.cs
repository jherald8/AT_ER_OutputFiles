using System;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Configuration;
using AT_ER_OutputFiles;
using System.IO;

namespace AT_ER_OutputFiles
{
    class Program
    {
        static void Main(string[] args)
        {
            string processType = ConfigurationSettings.AppSettings["ProcessType"];
            FileTool fileTool = new FileTool();
            ExecuteScript executePython = new ExecuteScript();
            if (processType == "1")
            {
                AT_ER_OutputFile outputFiles_Comparison = new AT_ER_OutputFile();
                outputFiles_Comparison.ProcessOfFiles();
            }
            else if (processType == "2")
            {
                AT_Encoding_Types outputFiles_Encoding = new AT_Encoding_Types();
                outputFiles_Encoding.Process();
            }
            else if (processType == "3")
            {
                string source = ConfigurationSettings.AppSettings["SourcePath"];
                string[] files = Directory.GetFiles(source);
                foreach (var file in files)
                    if (Path.GetExtension(file) == ".zip" || Path.GetExtension(file) == ".ZIP")
                    {
                        fileTool.Decompress(source);
                        break;
                    }
                files = Directory.GetFiles(source);
                foreach (var file in files)
                    fileTool.Decrypting(file);
            }
            else if (processType == "4")
            {
                executePython.DownloadGmail();
            }
            else if (processType == "5")
            {
                fileTool.SFTPConnect();
            }
        }
    }
}
