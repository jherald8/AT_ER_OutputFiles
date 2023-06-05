using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ude;
using Excel = Microsoft.Office.Interop.Excel;

namespace AT_ER_OutputFiles
{
    internal class AT_Encoding_Types
    {
        string source = ConfigurationSettings.AppSettings["SourcePath"];
        bool isEncrypted = false;
        public void Process()
        {
            string[] newFiles = Directory.GetFiles(source);
            foreach (var file in newFiles)
            {
                if (Path.GetExtension(file) == ".zip" || Path.GetExtension(file) == ".ZIP")
                {
                    Decompress(source);
                    break;
                }
            }

            string logFile = $@"c:\temp\EncodingLog-{DateTime.Now.ToString("MM-d-yy_HH-mm-ss")}.txt";
            StreamWriter sw;
            sw = File.CreateText(logFile);
            foreach (var file in newFiles)
            {
                
                byte[] fileBytes = File.ReadAllBytes(file);

                if (fileBytes.Length >= 3 && fileBytes[0] == 0xEF && fileBytes[1] == 0xBB && fileBytes[2] == 0xBF)
                {
                    // BOM indicates UTF-8

                    sw.WriteLine($"{Path.GetFileName(file)} : UTF-8 BOM");
                }
                else if (fileBytes.Length >= 2 && fileBytes[0] == 0xFF && fileBytes[1] == 0xFE)
                {
                    // BOM indicates UTF-16 Little Endian
                    sw.WriteLine($"{Path.GetFileName(file)} : UTF-16 LE BOM");
                }
                else
                {
                    // No BOM detected, use charset detection
                    using (var stream = new MemoryStream(fileBytes))
                    {
                        var detector = new CharsetDetector();
                        detector.Feed(stream);
                        detector.DataEnd();
                        string charset = detector.Charset;
                        Encoding.GetEncoding(charset);
                        sw.WriteLine($"{Path.GetFileName(file)} : UTF-8");
                    }
                }
            }
            sw.Close();
        }

        #region Unzip & Decrypt
        public void Decompress(string zip)
        {
            string temp = Path.Combine(source + @"Temp\");
            string[] files = Directory.GetFiles(zip, "*.zip", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                ZipFile.ExtractToDirectory(file, temp);
                File.Delete(file);
                string path = Path.GetDirectoryName(file);
                string[] tempFiles = Directory.GetFiles(temp);
                foreach (var fileTemp in tempFiles)
                {
                    int fileIncrement = 1;
                    string fileName = Path.GetFileName(fileTemp);
                    string fullPath = Path.Combine(path, fileName);
                    string withoutExt = Path.GetFileNameWithoutExtension(fileTemp);
                    withoutExt = withoutExt + "-com";
                    string itsExt = Path.GetExtension(fileTemp);
                    if (!File.Exists(fullPath))
                        File.Move(fileTemp, source + withoutExt + itsExt);
                    else
                    {
                        string newPathFile = path + "\\" + withoutExt + $"({fileIncrement}){itsExt}";
                        while (File.Exists(newPathFile))
                            newPathFile = path + "\\" + withoutExt + $"({fileIncrement++}){itsExt}";
                        File.Move(fileTemp, newPathFile);
                    }
                }
            }
            Directory.Delete(temp, true);
        }
        public string Decrypting(string newFiles)
        {
            string fileName = null;
            string extension = null;
            if (string.Equals((Path.GetExtension(newFiles)), ".txt", StringComparison.OrdinalIgnoreCase) /** ||
                string.Equals((Path.GetExtension(newFiles)), ".xml", StringComparison.OrdinalIgnoreCase) **/)
            {
                if (string.Equals((Path.GetExtension(newFiles)), ".txt", StringComparison.OrdinalIgnoreCase))
                    extension = ".txt";
                //if (string.Equals((Path.GetExtension(newFiles)), ".xml", StringComparison.OrdinalIgnoreCase))
                //    extension = ".xml";
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
                //if (string.Equals((Path.GetExtension(newFiles)), ".pdf", StringComparison.OrdinalIgnoreCase))
                //    extension = ".pdf";

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
        public void EncryptedChecker(string fileOne)
        {
            if (string.Equals((Path.GetExtension(fileOne)), ".txt", StringComparison.OrdinalIgnoreCase))
            {
                string[] lines = File.ReadAllLines(fileOne);
                foreach (string line in lines)
                {
                    if (line.Contains("AES"))
                    {
                        isEncrypted = true;
                        break;
                    }
                }
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xls", StringComparison.OrdinalIgnoreCase))
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileOne);
                Excel.Application xlAppTwo = new Excel.Application();
                Excel._Worksheet xlWorksheet;
                xlWorksheet = xlWorkbook.Sheets[1];

                for (int i = 1; i <= xlWorksheet.UsedRange.Rows.Count; i++) // row {1}
                {
                    for (int j = 1; j <= xlWorksheet.UsedRange.Columns.Count; j++) // col {A}
                    {
                        if (!String.IsNullOrEmpty(xlWorksheet.Cells[i, j].Text.ToString()) || !String.IsNullOrWhiteSpace(xlWorksheet.Cells[i, j].Text.ToString())) { }
                        if (xlWorksheet.Cells[i, j].Text.ToString().Contains("AES"))
                        {
                            isEncrypted = true;
                            break;
                        }
                    }
                }
                xlApp.Workbooks.Close();
                xlApp.Quit();
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    FileInfo fiOne = new FileInfo(fileOne);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelPackage excelOne = new ExcelPackage(fiOne);
                }
                catch (Exception)
                {
                    isEncrypted = true;
                }
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".csv", StringComparison.OrdinalIgnoreCase))
            {
                using (StreamReader f1 = new StreamReader(fileOne))
                {
                    var line1 = f1.ReadLine();
                    if (line1.Contains("AES"))
                        isEncrypted = true;
                }
            }
            //else if (string.Equals((Path.GetExtension(fileOne)), ".pdf", StringComparison.OrdinalIgnoreCase))
            //{
            //    try
            //    {
            //        iText.PdfReader pdfOne = new iText.PdfReader(fileOne);
            //    }
            //    catch (Exception)
            //    {
            //        isEncrypted = true;
            //    }
            //}
            //else if (string.Equals((Path.GetExtension(fileOne)), ".xml", StringComparison.OrdinalIgnoreCase))
            //{
            //    string[] lines = File.ReadAllLines(fileOne);
            //    foreach (string line in lines)
            //    {
            //        if (line.Contains("AES"))
            //        {
            //            isEncrypted = true;
            //            break;
            //        }
            //    }
            //}
        }
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
        #endregion
    }
}
