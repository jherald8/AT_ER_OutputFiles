using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using SharpAESCrypt;

namespace QATest
{
    internal class AESDecrypt
    {
        public void Decrypting()
        {
            Unzip unzip = new Unzip();
            unzip.Decompress();
            int count = 0;
            string[] files = Directory.GetFiles(@"C:\Users\jmartin\Downloads\Jerald Files\DailyTask Test\QATest\decryptFiles\", "*txt");
            
            foreach (var file in files)
            {
                //File.Move(file, Path.ChangeExtension(file, ".txt.aes"));
                File.Copy(file, Path.ChangeExtension(file, ".aes"));
                string test1 = Path.GetFileNameWithoutExtension(file);
                using (var outputfile = System.IO.File.OpenWrite(file))
                {
                    using (var inputfile = System.IO.File.OpenRead(@"C:\Users\jmartin\Downloads\Jerald Files\DailyTask Test\QATest\decryptFiles\" + test1 + ".aes"))
                    using (var encStream = new SharpAESCrypt.SharpAESCrypt("1", inputfile, SharpAESCrypt.OperationMode.Decrypt))
                    {
                        encStream.CopyTo(outputfile);
                    }
                }
                string test11 = file.Remove(file.Length - 4);
                File.Delete(test11 + ".aes");
            }
        }

    }
}
