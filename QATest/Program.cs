using System;
using System.Security.Cryptography;
using System.Runtime.InteropServices;

namespace QATest
{
    class Program
    {
        static void Main(string[] args)
        {
            AESDecrypt aesDecrypt = new AESDecrypt();
            aesDecrypt.Decrypting();

            //Unzip unzip = new Unzip();
            //unzip.Decompress();
            //TextComparison textComparison = new TextComparison();
            //textComparison.CompareTxtFiles();
            //ExcelComparison excelComparison = new ExcelComparison();
            //excelComparison.CompareExcelFiles();
            //CSVComparison csvComparison = new CSVComparison();
            //csvComparison.CompareCSVFiles();
        }
    }
}
