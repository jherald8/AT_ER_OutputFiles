using System;
using System.Security.Cryptography;
using System.Runtime.InteropServices;

namespace QATest
{
    class Program
    {
        static void Main(string[] args)
        {
            //StartProcess startProcess = new StartProcess();
            //startProcess.FileAutomation();

            OutputFiles_Comparison outputFiles_Comparison = new OutputFiles_Comparison();
            outputFiles_Comparison.ProcessOfFiles();
        }
    }
}
