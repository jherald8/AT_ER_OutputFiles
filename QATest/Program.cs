using System;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Configuration;
using AT_ER_OutputFiles;
using System.IO;

namespace QATest
{
    class Program
    {
        static void Main(string[] args)
        {
            string source = ConfigurationSettings.AppSettings["SourcePath"];
            string processType = ConfigurationSettings.AppSettings["ProcessType"];
            if (processType == "1")
            {
                OutputFiles_Comparison outputFiles_Comparison = new OutputFiles_Comparison();
                outputFiles_Comparison.ProcessOfFiles();
            }
            else if (processType == "2")
            {
                AT_Encoding_Types outputFiles_Encoding = new AT_Encoding_Types();
                outputFiles_Encoding.Process();
            }
            else if (processType == "3")
            {
                FileTool fileTool = new FileTool();
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
        }
    }
}
