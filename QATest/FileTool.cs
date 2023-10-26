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
using iText = iTextSharp.text.pdf;
using iTextParser = iTextSharp.text.pdf.parser;

namespace AT_ER_OutputFiles
{
    internal class FileTool
    {
        string pType = ConfigurationSettings.AppSettings["ProcessType"];
        string source = ConfigurationSettings.AppSettings["SourcePath"];
        public bool isEncrypted = false;

        #region Unzip & Decrypt
        public void ProcessTypeChanger()
        {
            if (pType == "2")
                source = ConfigurationSettings.AppSettings["EncodingPath"];
            else
                Console.Write("Do nothing!");

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
        public void Decompress(string zip)
        {
            ProcessTypeChanger();
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
        public void EncryptedChecker(string fileOne)
        {
            isEncrypted = false;
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
                CloseExcel();
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    FileInfo fiOne = new FileInfo(fileOne);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelPackage excelOne = new ExcelPackage(fiOne);
                    var wsOne = excelOne.Workbook.Worksheets[0];
                    int row = wsOne.Dimension.End.Row;
                    int col = wsOne.Dimension.End.Column;
                    for (int i = 1; i <= row; i++)
                    {
                        for (int j = 1; j <= col; j++)
                        {
                            if (wsOne.Cells[i, j].Text.Contains("AES"))
                                isEncrypted = true;
                        }
                    }
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
            else if (string.Equals((Path.GetExtension(fileOne)), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    iText.PdfReader pdfOne = new iText.PdfReader(fileOne);
                    int pageFileOne = pdfOne.NumberOfPages;
                    string dataOne;
                    for (int i = 1; i <= pageFileOne; i++)
                    {
                        if (i <= 1)
                        {
                            dataOne = iTextParser.PdfTextExtractor.GetTextFromPage(pdfOne, i, new iTextParser.LocationTextExtractionStrategy());
                            if (dataOne.Contains("AES"))
                                isEncrypted = true;
                        }
                    }
                }
                catch (Exception)
                {
                    isEncrypted = true;
                }
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xml", StringComparison.OrdinalIgnoreCase))
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
        }
        public string Decrypting(string newFiles)
        {
            ProcessTypeChanger();
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
