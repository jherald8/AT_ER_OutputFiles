using System;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Configuration;
using AT_ER_OutputFiles;

namespace QATest
{
    class Program
    {
        static void Main(string[] args)
        {
            string processType = ConfigurationSettings.AppSettings["ProcessType"];
            if (processType == "1")
            {
                OutputFiles_Comparison outputFiles_Comparison = new OutputFiles_Comparison();
                outputFiles_Comparison.ProcessOfFiles();
            }
            else if( processType == "2")
            {
                AT_Encoding_Types outputFiles_Encoding = new AT_Encoding_Types();
                outputFiles_Encoding.Process();
            }
        }
    }
}
