using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using System.IO;

namespace QATest
{
    internal class Unzip
    {
        public void Decompress()
        {
            string[] zipFilePath = Directory.GetFiles(@"C:\Users\jmartin\Downloads\Jerald Files\DailyTask Test\QATest\decryptFiles", "*zip");
            string extractionPath = @"C:\Users\jmartin\Downloads\Jerald Files\DailyTask Test\QATest\decryptFiles\";
            foreach (var file in zipFilePath)
            {
                ZipFile.ExtractToDirectory(file, extractionPath);
                Console.WriteLine("Extracted Successfully");
            }
            //TextComparison textComparison = new TextComparison();
            //textComparison.CompareTxtFiles();
        }
    }
}
