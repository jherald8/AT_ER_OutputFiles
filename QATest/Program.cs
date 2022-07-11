using System;

namespace QATest
{
    class Program
    {
        static void Main(string[] args)
        {
            AESDecrypt aesDecrypt = new AESDecrypt();
            aesDecrypt.Decrypt("1", @"C:\Users\jmartin\Downloads\Jerald Files\DailyTask Test\QATest\decryptFiles\OPSH-555DEL20220711172444.txt.aes", @"C:\Users\jmartin\Downloads\Jerald Files\DailyTask Test\QATest\decryptFiles\OPSH-555DEL202207111724441.txt.aes");
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
