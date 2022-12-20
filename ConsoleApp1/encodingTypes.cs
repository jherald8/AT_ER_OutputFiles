using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QATestEncoding
{
    internal class encodingTypes
    {
        public Encoding GetEncoding(string filename = @"C:\Users\jmartin\Downloads\OPSH-9322TXT20220826191725.txt")
        {
            using (FileStream stream = File.OpenRead(filename))
            {
                Encoding enc = Encoding.UTF8;

                stream.Seek(0, SeekOrigin.Begin);
                Ude.ICharsetDetector cdet = new Ude.CharsetDetector();

                cdet.Feed(stream);
                cdet.DataEnd();

                if (cdet.Charset != null)
                {
                    enc = Encoding.GetEncoding(cdet.Charset);

                }
                string test = enc.BodyName.ToString();
                return enc;
            }
            
        }
        //public void Encoding()
        //{
        //    var filename = @"C:\Users\jmartin\Downloads\OPSH-9322TXT20220826191725.txt";
        //    string typeEncoding = "";
        //    var bom = new byte[4];
        //    using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read))
        //    {
        //        file.Read(bom, 0, 4);
        //    }

        //    // Analyze the BOM
        //    if (bom[0] == 0x2b && bom[1] == 0x2f && bom[2] == 0x76) 
        //        typeEncoding = "UTF7";
        //    if (bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf) 
        //        typeEncoding = "UTF8";
        //    if (bom[0] == 0xff && bom[1] == 0xfe && bom[2] == 0 && bom[3] == 0) 
        //        typeEncoding = "UTF32"; //UTF-32LE
        //    if (bom[0] == 0xff && bom[1] == 0xfe) 
        //        typeEncoding = "UTF16LE"; //UTF-16LE
        //    if (bom[0] == 0xfe && bom[1] == 0xff) 
        //        typeEncoding = "UTF16BE"; //UTF-16BE
        //    if (bom[0] == 0 && bom[1] == 0 && bom[2] == 0xfe && bom[3] == 0xff)
        //        typeEncoding = "UTF32BE";  //UTF-32BE

        //    // We actually have no idea what the encoding is if we reach this point, so
        //    // you may wish to return null instead of defaulting to ASCII

        //    Console.Write(typeEncoding);
        //}
        public void DecryptingExcel(string source)
        {
            //string[] files = Directory.GetFiles(source, "*xlsx");
            //foreach (var file in files)
            //{
            //    File.Copy(file, Path.ChangeExtension(file, ".aes"));
            //    string fileName = Path.GetFileNameWithoutExtension(file);
            //    var temp = Directory.CreateDirectory(source + @"temp\");
            //    using (var outputfile = System.IO.File.OpenWrite(temp + fileName + ".xlsx"))
            //    {
            //        using (var inputfile = System.IO.File.OpenRead(source + fileName + ".aes"))
            //        using (var encStream = new SharpAESCrypt.SharpAESCrypt("1", inputfile, SharpAESCrypt.OperationMode.Decrypt))
            //        {
            //            encStream.CopyTo(outputfile);
            //        }
            //    }
            //    File.Delete(file);
            //    File.Delete(source + fileName + ".aes");
            //    File.Copy(temp + fileName + ".xlsx", source + fileName + ".xlsx");
            //    Directory.Delete(source + @"temp\", true);

            //string[] files = Directory.GetFiles(source, "*csv");
            //foreach (var file in files)
            //{
            //    File.Copy(file, Path.ChangeExtension(file, ".aes"));
            //    string fileName = Path.GetFileNameWithoutExtension(file);
            //    var temp = Directory.CreateDirectory(source + @"temp\");
            //    using (var outputfile = System.IO.File.OpenWrite(temp + fileName + ".csv"))
            //    {
            //        using (var inputfile = System.IO.File.OpenRead(source + fileName + ".aes"))
            //        using (var encStream = new SharpAESCrypt.SharpAESCrypt("1", inputfile, SharpAESCrypt.OperationMode.Decrypt))
            //        {
            //            encStream.CopyTo(outputfile);
            //        }
            //    }
            //    File.Delete(file);
            //    File.Delete(source + fileName + ".aes");
            //    File.Copy(temp + fileName + ".csv", source + fileName + ".csv");
            //    Directory.Delete(source + @"temp\", true);

            //    //string FinalText = System.Text.Encoding.UTF8.GetString(byteArray);

            //    //string remExt = file.Remove(file.Length - 4);
            //    //File.Delete(remExt + ".aes");
            //}

            string[] files = Directory.GetFiles(source, "*xls");
            foreach (var file in files)
            {
                File.Copy(file, Path.ChangeExtension(file, ".aes"));
                string fileName = Path.GetFileNameWithoutExtension(file);
                var temp = Directory.CreateDirectory(source + @"temp\");
                using (var outputfile = System.IO.File.OpenWrite(temp + fileName + ".xls"))
                {
                    using (var inputfile = System.IO.File.OpenRead(source + fileName + ".aes"))
                    using (var encStream = new SharpAESCrypt.SharpAESCrypt("1", inputfile, SharpAESCrypt.OperationMode.Decrypt))
                    {
                        encStream.CopyTo(outputfile);
                    }
                }
                File.Delete(file);
                File.Delete(source + fileName + ".aes");
                File.Copy(temp + fileName + ".xls", source + fileName + ".xls");
                Directory.Delete(source + @"temp\", true);

                //string FinalText = System.Text.Encoding.UTF8.GetString(byteArray);

                //string remExt = file.Remove(file.Length - 4);
                //File.Delete(remExt + ".aes");
            }

            //string[] aesFiles = Directory.GetFiles(source, "*xlsx");
            //foreach (var file in aesFiles)
            //{
            //    string[] lines = File.ReadAllLines(file);
            //    List<string> fixedLines = new List<string>();
            //    int countLine = 0;
            //    foreach (var line in lines)
            //    {
            //        countLine++;
            //        if (countLine <= 9)
            //        {
            //            fixedLines.Add(line);
            //        }
            //    }
            //    File.Delete(file);
            //    System.IO.File.WriteAllLines(file, fixedLines);
            //}
        }
    }
}
