﻿using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using iText = iTextSharp.text.pdf;
using iTextParser = iTextSharp.text.pdf.parser;

namespace QATest
{
    internal class OutputFiles_Comparison
    {
        string source = ConfigurationSettings.AppSettings["SourcePath"];
        public void ProcessOfFiles()
        {
            string destination = ConfigurationSettings.AppSettings["DestinationPath"];

            string[] newFiles = Directory.GetFiles(source);
            string[] oldFiles = Directory.GetFiles(destination);

            foreach (var file in newFiles)
            {
                if (Path.GetExtension(file) == ".zip" || Path.GetExtension(file) == ".ZIP" )
                {
                    Decompress(source);
                    break;
                }
            }
            newFiles = Directory.GetFiles(source);

            string passedPath = ConfigurationSettings.AppSettings["PassedPath"];
            string failedPath = ConfigurationSettings.AppSettings["FailedPath"];
            bool encrypted = false;
            StartProcess startProcess = new StartProcess();
            
            foreach (var fileOne in newFiles)
            {
                encrypted = false;
                foreach (var fileTwo in oldFiles)
                {
                    string subFileOne = Regex.Match(Path.GetFileNameWithoutExtension(fileOne), "[a-zA-Z]+-[0-9]{1,10}[a-zA-Z]{1,7}").ToString();
                    string subFileTwo = Regex.Match(Path.GetFileNameWithoutExtension(fileTwo), "[a-zA-Z]+-[0-9]{1,10}[a-zA-Z]{1,7}").ToString();
                    if (subFileOne == subFileTwo)
                    {
                        if (Path.GetExtension(fileOne) == ".txt" || Path.GetExtension(fileOne) == ".TXT")
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
                        else if (Path.GetExtension(fileOne) == ".xls" || Path.GetExtension(fileOne) == ".XLS")
                        {
                            try
                            {
                                CompareXLSFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                            catch (Exception)
                            {
                                CloseExcel();
                                Decrypting(fileOne);
                                CloseExcel();
                                CompareXLSFiles(fileOne, fileTwo, passedPath, failedPath);
                                CloseExcel();
                            }
                        }
                        else if (Path.GetExtension(fileOne) == ".xlsx" || Path.GetExtension(fileOne) == ".XLSX")
                        {
                            try
                            {
                                CompareXLSXFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                            catch (Exception)
                            {
                                Decrypting(fileOne);
                                CompareXLSXFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                        }
                        else if (Path.GetExtension(fileOne) == ".csv" || Path.GetExtension(fileOne) == ".CSV")
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
                        else if (Path.GetExtension(fileOne) == ".pdf" || Path.GetExtension(fileOne) == ".PDF")
                        {
                            try
                            {
                                ComparePDFFiles(fileOne, fileTwo, passedPath, failedPath);
                            }
                            catch (Exception)
                            {
                                Decrypting(fileOne);
                                ComparePDFFiles(fileOne, fileTwo, passedPath, failedPath);
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
            bool cancel = false;
            for (int i = 1; i <= row; i++)
            {
                for (int j = 1; j <= col; j++)
                {
                    if (!String.IsNullOrEmpty(xlWorksheet.Cells[i, j].Text.ToString()) || !String.IsNullOrWhiteSpace(xlWorksheet.Cells[i, j].Text.ToString()))
                        if (xlWorksheet.Cells[i, j].Text.ToString().Contains("AES"))
                            throw new Exception("lalala");

                    if (!String.IsNullOrEmpty(xlWorksheet.Cells[i, j].Text.ToString()) || !String.IsNullOrWhiteSpace(xlWorksheet.Cells[i, j].Text.ToString()))
                        if (xlWorksheet.Cells[i, j].Text.ToString() != xlWorksheetTwo.Cells[i, j].Text.ToString())
                        {
                            string getFileName = Path.GetFileName(fileOne);
                            File.Copy(fileOne, failedPath + Path.GetFileName(fileOne));
                            cancel = true;
                            break;
                        }   
                }
                if (cancel)
                    break;
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
                Console.WriteLine("Passed");
            else
            {
                File.Copy(fileOne, failedPath + Path.GetFileName("/" + fileOne));
            }
        }
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
                    string itsExt = Path.GetExtension(fileTemp);
                    if (!File.Exists(fullPath))
                        File.Move(fileTemp, source + fileName);
                    else //
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
        public void Decrypting(string newFiles)
        {
            if (Path.GetExtension(newFiles) == ".txt" || Path.GetExtension(newFiles) == ".TXT")
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
                if (Path.GetExtension(newFiles) == ".xls" || Path.GetExtension(newFiles) == ".XLS")
                    extension = ".xls";
                if (Path.GetExtension(newFiles) == ".xlsx" || Path.GetExtension(newFiles) == ".XLSX")
                    extension = ".xlsx";
                if (Path.GetExtension(newFiles) == ".csv" || Path.GetExtension(newFiles) == ".CSV")
                    extension = ".csv";
                if (Path.GetExtension(newFiles) == ".pdf" || Path.GetExtension(newFiles) == ".PDF")
                    extension = ".pdf";

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
                Thread.Sleep(1500);
                if (extension == ".xls")
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
        public void CompareXLSXFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
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
                File.Copy(fileOne, failedPath + Path.GetFileName(fileTwo));
            else
                Console.WriteLine("Passed");

        }
        public void CompareCSVFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
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
                    if (line1.Contains("AES"))
                        throw new Exception("Your message here!");
                    if (line1 != line2)
                    {
                        File.Copy(fileOne, failedPath + Path.GetFileName(fileOne));
                        break;
                    }
                    else
                        Console.WriteLine("Passed");
                }
            }
        }
        public void ComparePDFFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
        {
            iText.PdfReader pdfOne = new iText.PdfReader(fileOne);
            iText.PdfReader pdfTwo = new iText.PdfReader(fileTwo);
            int pageFileOne = pdfOne.NumberOfPages;
            string[] wordOne;
            string[] wordTwo;
            string line;
            string dataOne;
            string dataTwo;
            string[] resultOne;
            string[] resultTwo;
            for (int i = 1; i <= pageFileOne; i++)
            {
                if(i <= 1)
                {
                    dataOne = iTextParser.PdfTextExtractor.GetTextFromPage(pdfOne, i, new iTextParser.LocationTextExtractionStrategy());
                    dataTwo = iTextParser.PdfTextExtractor.GetTextFromPage(pdfTwo, i, new iTextParser.LocationTextExtractionStrategy());
                    dataOne = Regex.Replace(dataOne, @"(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+\s-\s(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+", "");//remove date period
                    dataTwo = Regex.Replace(dataTwo, @"(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+\s-\s(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+", "");//remove page# 
                    if (Enumerable.SequenceEqual(dataOne, dataTwo) == true)
                        Console.WriteLine("Passed");
                    else
                        File.Copy(fileOne, failedPath + Path.GetFileName("/" + fileOne));
                }
            }
        }
    }
}
