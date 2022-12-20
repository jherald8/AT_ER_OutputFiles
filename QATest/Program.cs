using System;
using System.Security.Cryptography;
using System.Runtime.InteropServices;

namespace QATest
{
    class Program
    {
        static void Main(string[] args)
        {
            OutputFiles_Comparison outputFiles_Comparison = new OutputFiles_Comparison();
            outputFiles_Comparison.ProcessOfFiles();
        }
    }
}
