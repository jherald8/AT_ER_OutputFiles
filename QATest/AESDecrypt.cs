using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ExpressEncription;
using SharpAESCrypt;

namespace QATest
{
    internal class AESDecrypt
    {
        public static void Decrypt(string password, string inputfile, string outputfile)
        {
            //using (FileStream infs = File.OpenRead(inputfile))
            //using (FileStream outfs = File.Create(outputfile))
            //    Decrypt(password, infs, outfs);
        }

        public void Decrypting()
        {
            string[] files = Directory.GetFiles(@"C:\Users\jmartin\Downloads\Jerald Files\DailyTask Test\QATest\decryptFiles\", "*txt.aes");
            foreach (var file in files)
            {
                //File.Move(file, Path.ChangeExtension(file, ".txt.aes"));
                //ExpressEncription.AESEncription.AES_Decrypt(file, "1");
                
            }

        }

    }
}
