using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace QATest
{
    internal class OutputFiles_Comparison
    {
        string source = ConfigurationSettings.AppSettings["SourcePath"];
        public void ProcessOfFiles()
        {
            Decompress(source);

            string destination = ConfigurationSettings.AppSettings["DestinationPath"];

            string[] newFiles = Directory.GetFiles(source);
            string[] oldFiles = Directory.GetFiles(destination);

            string passedPath = ConfigurationSettings.AppSettings["PassedPath"];
            string failedPath = ConfigurationSettings.AppSettings["FailedPath"];
            StartProcess startProcess = new StartProcess();

            bool encrypted = false;
            foreach (var fileOne in newFiles)
            {
                encrypted = false;
                foreach (var fileTwo in oldFiles)
                {
                    //Compress
                    if (Path.GetFileNameWithoutExtension(fileOne) == Path.GetFileNameWithoutExtension(fileTwo))
                    {
                        if (Path.GetExtension(fileOne) == ".txt")
                        {
                            try
                            {
                                CompareTxtFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                            catch (Exception)
                            {
                                Decrypting(fileOne);
                                CompareTxtFiles(source, fileTwo, passedPath, failedPath);
                            }
                        }
                        else if (Path.GetExtension(fileOne) == ".xls")
                        {
                            try
                            {
                                CompareXLSFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                            catch (Exception)
                            {
                                CloseExcel();
                                Decrypting(fileOne);
                                CompareXLSFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                        }
                        else if (Path.GetExtension(fileOne) == ".xlsx")
                        {
                            try
                            {
                                CompareExcelFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                            catch (Exception)
                            {
                                Decrypting(fileOne);
                                CompareExcelFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                        }
                        else if (Path.GetExtension(fileOne) == ".csv")
                        {
                            try
                            {
                                CompareCSVFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                            catch (Exception)
                            {
                                Decrypting(fileOne);
                                CompareCSVFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                        }
                        source = ConfigurationSettings.AppSettings["SourcePath"];
                        break;
                    }
                }
            }
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
        public void CompareXLSFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileOne);

            Excel.Application xlAppTwo = new Excel.Application();
            Excel.Workbook xlWorkbookTwo = xlAppTwo.Workbooks.Open(fileTwo);

            Excel._Worksheet xlWorksheet;
            Excel._Worksheet xlWorksheetTwo;
            if (Path.GetFileName(fileOne).Contains("EXFMT") || Path.GetFileName(fileTwo).Contains("EXFMT"))
            {
                xlWorksheet = xlWorkbook.Sheets[2];
                xlWorksheetTwo = xlWorkbookTwo.Sheets[2];
            }
            else
            {
                xlWorksheet = xlWorkbook.Sheets[1];
                xlWorksheetTwo = xlWorkbookTwo.Sheets[1];
            }

            int row = xlWorksheet.UsedRange.Rows.Count;
            int col = xlWorksheet.UsedRange.Columns.Count;
            for (int i = 1; i <= row; i++)
            {
                for (int j = 1; j <= col; j++)
                {
                    try
                    {
                        if (!String.IsNullOrEmpty(xlWorksheet.Cells[i, j].Text.ToString()) || !String.IsNullOrWhiteSpace(xlWorksheet.Cells[i, j].Text.ToString()))
                            if (xlWorksheet.Cells[i, j].Text.ToString().Contains("AES"))
                                throw new Exception("lalala");

                        if (!String.IsNullOrEmpty(xlWorksheet.Cells[i, j].Text.ToString()) || !String.IsNullOrWhiteSpace(xlWorksheet.Cells[i, j].Text.ToString()))
                            if (xlWorksheet.Cells[i, j].Text.ToString() != xlWorksheetTwo.Cells[i, j].Text.ToString())
                                File.Copy(fileOne, failedPath + Path.GetFileName(fileOne));
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
            }
            xlApp.Workbooks.Close();
            xlApp.Quit();
            xlAppTwo.Workbooks.Close();
            xlAppTwo.Quit();
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlWorkbookTwo);
        }
        public void CompareTxtFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
        {
            if (Path.GetFileName(fileOne) == Path.GetFileName(fileTwo))
            {
                string[] lines = File.ReadAllLines(fileOne);
                string[] lines2 = File.ReadAllLines(fileTwo);

                List<string> newFiles = new List<string>();
                List<string> oldFiles = new List<string>();
                foreach (string line in lines)
                {
                    newFiles.Add(line);
                    if (line.Contains("AES"))
                        throw new Exception("Your message here!");
                }
                foreach (var line2 in lines2)
                {
                    oldFiles.Add(line2);
                }
                if (Enumerable.SequenceEqual(newFiles, oldFiles) == true)
                {
                    File.Copy(fileOne, passedPath + Path.GetFileName(fileOne));
                }
                else
                {
                    File.Copy(fileOne, failedPath + Path.GetFileName("/" + fileOne));
                }
            }
        }
        public void Decompress(string zip)
        {
            string temp = Path.Combine(source + @"Temp\");
            string[] files = Directory.GetFiles(zip);
            foreach (var file in files)
            {
                ZipFile.ExtractToDirectory(file, temp);
                File.Delete(file);
                string extConverter = Path.GetFileNameWithoutExtension(file) + ".txt";
                try
                {
                    File.Move(temp + extConverter, source + extConverter);
                }
                catch (Exception)
                {
                    string[] tempFiles = Directory.GetFiles(temp);
                    foreach (var fileTemp in tempFiles)
                    {
                        string fileName = Path.GetFileName(fileTemp);
                        File.Move(fileTemp, source + fileName);
                    }
                }
            }
            Directory.Delete(temp, true);
        }
        public void Decrypting(string newFiles)
        {
            if (Path.GetExtension(newFiles) == ".txt")
            {
                File.Copy(newFiles, Path.ChangeExtension(newFiles, ".aes"));
                string fileName = Path.GetFileNameWithoutExtension(newFiles);
                using (var outputfile = System.IO.File.OpenWrite(newFiles))
                {
                    using (var inputfile = System.IO.File.OpenRead(source + fileName + ".aes"))
                    using (var encStream = new SharpAESCrypt.SharpAESCrypt("1", inputfile, SharpAESCrypt.OperationMode.Decrypt))
                    {
                        encStream.CopyTo(outputfile);
                    }
                }
                string remExt = newFiles.Remove(newFiles.Length - 4);
                File.Delete(remExt + ".aes");

                string[] lines = File.ReadAllLines(newFiles);
                List<string> fixedLines = new List<string>();
                int countLine = 0;
                foreach (var line in lines)
                {
                    countLine++;
                    if (lines[0].Contains("Employee No"))
                    {
                        if (countLine <= 11)
                        {
                            fixedLines.Add(line);
                        }
                    }
                    else
                    {
                        if (countLine <= 10)
                        {
                            fixedLines.Add(line);
                        }
                    }
                }
                File.Delete(newFiles);
                System.IO.File.WriteAllLines(newFiles, fixedLines);
                source = source + Path.GetFileNameWithoutExtension(newFiles + ".txt");
            }
            else
            {
                string extension = null;
                if (Path.GetExtension(newFiles) == ".xls")
                    extension = ".xls";
                if (Path.GetExtension(newFiles) == ".xlsx")
                    extension = ".xlsx";
                if (Path.GetExtension(newFiles) == ".csv")
                    extension = ".csv";

                File.Copy(newFiles, Path.ChangeExtension(newFiles, ".aes"));
                string fileName = Path.GetFileNameWithoutExtension(newFiles);
                var temp = Directory.CreateDirectory(source + @"temp\");
                using (var outputfile = System.IO.File.OpenWrite(temp + fileName + extension))
                {
                    using (var inputfile = System.IO.File.OpenRead(source + fileName + ".aes"))
                    using (var encStream = new SharpAESCrypt.SharpAESCrypt("1", inputfile, SharpAESCrypt.OperationMode.Decrypt))
                    {
                        encStream.CopyTo(outputfile);
                    }
                }
                if(extension == ".xls")
                    CloseExcel();
                File.Delete(newFiles);
                File.Delete(source + fileName + ".aes");
                File.Copy(temp + fileName + extension, source + fileName + extension);
                Directory.Delete(source + @"temp\", true);
            }
        }
        public void CompareFileName(string[] fileOne, string[] fileTwo, string passedPath, string failedPath)
        {
            int oneCounter = 0;

            foreach (var oneFile in fileTwo)
            {
                oneCounter++;
                int twoCounter = 1;


                foreach (var twoFile in fileOne)
                {
                    if (Path.GetFileName(oneFile.Remove(oneFile.Length - 18)) == Path.GetFileName(twoFile.Remove(twoFile.Length - 18))
                        && oneCounter == twoCounter)
                    {
                        File.Copy(oneFile, passedPath + Path.GetFileName("/" + oneFile));
                        break;
                    }
                    else if (Path.GetFileName(oneFile.Remove(oneFile.Length - 18)) != Path.GetFileName(twoFile.Remove(twoFile.Length - 18))
                        && oneCounter == twoCounter)
                    {
                        File.Copy(oneFile, failedPath + Path.GetFileName("/" + oneFile));
                        break;
                    }
                    twoCounter++;
                }
            }
        }
        public void CompareExcelFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
        {
            if (Path.GetFileName(fileOne) == Path.GetFileName(fileTwo))
            {
                FileInfo fiOne = new FileInfo(fileOne);
                FileInfo fiTwo = new FileInfo(fileTwo);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage excelOne = new ExcelPackage(fiOne);
                ExcelPackage excelTwo = new ExcelPackage(fiTwo);
                var wsOne = excelOne.Workbook.Worksheets[0];
                var wsTwo = excelTwo.Workbook.Worksheets[0];
                int row = wsOne.Dimension.End.Row;
                int col = wsOne.Dimension.End.Column;
                bool isFailed = false;
                for (int i = 1; i <= row; i++)
                {
                    for (int j = 1; j <= col; j++)
                    {
                        if (wsOne.Cells[i, j].Text.Contains("AES"))
                            throw new Exception("lalala");
                        if (wsOne.Cells[i, j].Text != wsTwo.Cells[i, j].Text)
                        {
                            using (ExcelRange rng = wsOne.Cells[i, j])
                            {
                                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                            }
                            isFailed = true;
                        }
                    }
                }
                excelOne.Save();
                if (isFailed)
                {
                    File.Move(fileOne, failedPath + Path.GetFileName(fileTwo));
                }
            }
        }
        public void CompareCSVFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
        {
            if (Path.GetFileName(fileOne) == Path.GetFileName(fileTwo))
            {
                using (StreamReader f1 = new StreamReader(fileOne))
                using (StreamReader f2 = new StreamReader(fileTwo))
                {
                    var differences = new List<string>();
                    int lineNumber = 0;
                    while (!f1.EndOfStream)
                    {
                        if (f2.EndOfStream)
                        {
                            differences.Add("Differing number of lines - f2 has less.");
                            break;
                        }

                        lineNumber++;
                        var line1 = f1.ReadLine();
                        var line2 = f2.ReadLine();
                        if(line1.Contains("AES"))
                            throw new Exception("Your message here!");
                        if (line1 != line2)
                        {
                            File.Copy(fileOne, failedPath + Path.GetFileName(fileOne));
                            break;
                        }
                    }
                }
            }
        }
    }
}
