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
            string[] zipFilePath = Directory.GetFiles(@"D:\Work\TestQA\CompressedFile\sourceFile\", "*zip");
            string extractionPath = @"D:\Work\TestQA\CompressedFile\destFile\";
            foreach (var file in zipFilePath)
            {
                ZipFile.ExtractToDirectory(file, extractionPath);
                File.Delete(file);
            }
        }
    }
}
