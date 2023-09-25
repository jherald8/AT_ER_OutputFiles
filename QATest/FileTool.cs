using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AT_ER_OutputFiles
{
    internal class FileTool
    {
        string source = ConfigurationSettings.AppSettings["SourcePath"];
        bool isEncrypted = false;

        #region Unzip & Decrypt
        private static void CloseExcel(Excel.Application ExcelApplication = null)
        {
            if (ExcelApplication != null)
            {
                ExcelApplication.Workbooks.Close();
                ExcelApplication.Quit();
            }

            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.MainWindowTitle.Length == 0) { PK.Kill(); }
            }
        }
        public void Decompress(string zip)
        {
            string temp = Path.Combine(source + @"Temp\");
            string[] files = Directory.GetFiles(zip, "*.zip", SearchOption.TopDirectoryOnly);
            foreach (var file in files)
            {
                bool isLower = false;
                if (Path.GetFileNameWithoutExtension(file).Any(char.IsLower))
                    isLower = true;
                string newPathFile = null;
                ZipFile.ExtractToDirectory(file, temp);
                File.Delete(file);
                string[] tempFiles = Directory.GetFiles(temp);
                foreach (var fileTemp in tempFiles)
                {
                    int fileIncrement = 1;
                    newPathFile = Path.Combine(source, Path.GetFileNameWithoutExtension(fileTemp) + Path.GetExtension(fileTemp));
                    if (!File.Exists(newPathFile)) //notExist
                    {
                        if (isLower == true)
                        {
                            newPathFile = newPathFile.Replace("-com", "");
                            File.Move(fileTemp, newPathFile.Replace("-com", "")); //no com
                        }
                        else if (isLower == false)
                        {
                            int getIndex = newPathFile.IndexOf(Path.GetExtension(newPathFile));
                            newPathFile = newPathFile.Insert(getIndex, "-com");
                            while (File.Exists(newPathFile))
                            {
                                newPathFile = source + Path.GetFileNameWithoutExtension(fileTemp) + $"({fileIncrement++})-com{Path.GetExtension(fileTemp)}";
                            }
                            File.Move(fileTemp, newPathFile); //add com
                        }

                    }
                    else if (File.Exists(newPathFile)) //isExist
                    {
                        newPathFile = source + Path.GetFileNameWithoutExtension(fileTemp) + $"{Path.GetExtension(fileTemp)}";
                        if (isLower == true) //no com
                        {
                            newPathFile = newPathFile.Replace("-com", "");
                            while (File.Exists(newPathFile))
                            {
                                newPathFile = source + Path.GetFileNameWithoutExtension(fileTemp) + $"({fileIncrement++}){Path.GetExtension(fileTemp)}";
                            }
                            File.Move(fileTemp, newPathFile);
                        }
                        else if (isLower == false) //add com
                        {
                            int getIndex = newPathFile.IndexOf(Path.GetExtension(newPathFile));
                            newPathFile = newPathFile.Insert(getIndex, "-com");
                            while (File.Exists(newPathFile))
                            {
                                newPathFile = source + Path.GetFileNameWithoutExtension(fileTemp) + $"({fileIncrement++})-com{Path.GetExtension(fileTemp)}";
                            }
                            File.Move(fileTemp, newPathFile);
                        }
                    }
                }
            }
            Directory.Delete(temp, true);
        }
        public string Decrypting(string newFiles)
        {
            string fileName = null;
            string extension = null;
            if (string.Equals((Path.GetExtension(newFiles)), ".txt", StringComparison.OrdinalIgnoreCase) ||
                string.Equals((Path.GetExtension(newFiles)), ".xml", StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals((Path.GetExtension(newFiles)), ".txt", StringComparison.OrdinalIgnoreCase))
                    extension = ".txt";
                if (string.Equals((Path.GetExtension(newFiles)), ".xml", StringComparison.OrdinalIgnoreCase))
                    extension = ".xml";
                File.Copy(newFiles, Path.ChangeExtension(newFiles, ".aes"));
                fileName = Path.GetFileNameWithoutExtension(newFiles);
                var temp = Directory.CreateDirectory(source + @"temp\");
                using (var outputfile = System.IO.File.OpenWrite(temp + fileName + extension))
                {
                    using (var inputfile = System.IO.File.OpenRead(source + fileName + ".aes"))
                    using (var encStream = new SharpAESCrypt.SharpAESCrypt("1", inputfile, SharpAESCrypt.OperationMode.Decrypt))
                    {
                        encStream.CopyTo(outputfile);
                    }
                }
                File.Delete(newFiles);
                File.Delete(source + fileName + ".aes");
                File.Copy(temp + fileName + extension, source + fileName + "-enc" + extension);
                Directory.Delete(source + @"temp\", true);
                //source = source + Path.GetFileNameWithoutExtension(newFiles + ".txt");
            }
            else
            {
                if (string.Equals((Path.GetExtension(newFiles)), ".xls", StringComparison.OrdinalIgnoreCase))
                    extension = ".xls";
                if (string.Equals((Path.GetExtension(newFiles)), ".xlsx", StringComparison.OrdinalIgnoreCase))
                    extension = ".xlsx";
                if (string.Equals((Path.GetExtension(newFiles)), ".csv", StringComparison.OrdinalIgnoreCase))
                    extension = ".csv";
                if (string.Equals((Path.GetExtension(newFiles)), ".pdf", StringComparison.OrdinalIgnoreCase))
                    extension = ".pdf";

                File.Copy(newFiles, Path.ChangeExtension(newFiles, ".aes"));
                fileName = Path.GetFileNameWithoutExtension(newFiles);
                var temp = Directory.CreateDirectory(source + @"temp\");
                using (var outputfile = System.IO.File.OpenWrite(temp + fileName + extension))
                {
                    using (var inputfile = System.IO.File.OpenRead(source + fileName + ".aes"))
                    using (var encStream = new SharpAESCrypt.SharpAESCrypt("1", inputfile, SharpAESCrypt.OperationMode.Decrypt))
                    {
                        encStream.CopyTo(outputfile);
                    }
                }
                if (extension == ".xls")
                    CloseExcel();
                File.Delete(newFiles);
                File.Delete(source + fileName + ".aes");
                File.Copy(temp + fileName + extension, source + fileName + "-enc" + extension);
                Directory.Delete(source + @"temp\", true);
            }
            newFiles = source + fileName + "-enc" + extension;
            return newFiles;
        }
        #endregion
    }
}
