﻿using OfficeOpenXml;
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
using Renci.SshNet;


namespace AT_ER_OutputFiles
{
    internal class OutputFiles_Comparison
    {
        #region Declarations
        string source = ConfigurationSettings.AppSettings["SourcePath"]; //new
        bool isPassed = false;
        bool isEncrypted = false;
        #endregion
        #region Path creator
        //public void SFTPConnect()
        //{
        //    string server = "52.221.171.219";
        //    string username = "spintest";
        //    string password = "Spinifex01!";
        //    int port = 22;
        //    using (var client = new SshClient(server, port, username, password))
        //    {
        //        client.Connect();
        //    }
        //}
        public void PathCreator()
        {
            string destination = @"c:\temp\";
            if (!Directory.Exists(destination))
                Directory.CreateDirectory(destination);
            string oldFiles = Path.Combine(destination, "oldfiles");
            string newFiles = Path.Combine(destination, "newfiles");
            string failedPath = Path.Combine(destination, "failedpath");
            if (!Directory.Exists(oldFiles))
                Directory.CreateDirectory(oldFiles);
            if (!Directory.Exists(newFiles))
                Directory.CreateDirectory(newFiles);
            if (!Directory.Exists(failedPath))
                Directory.CreateDirectory(failedPath);
        }
        #endregion
        #region Process of Files
        public void ProcessOfFiles()
        {
            string pathCreator = ConfigurationSettings.AppSettings["PathCreator"].ToLower();
            #region PathCreator 
            if (pathCreator == "on")
            {
                PathCreator();
            }
            #endregion
            string destination = ConfigurationSettings.AppSettings["DestinationPath"];
            string failedPath = ConfigurationSettings.AppSettings["FailedPath"];
            string decrypt = ConfigurationSettings.AppSettings["Decryption"];
            
            string[] newFiles = Directory.GetFiles(source);
            string[] oldFiles = Directory.GetFiles(destination);

            

            #region Decompress
            string logFile = $@"c:\temp\LOG-{DateTime.Now.ToString("MM-d-yy-HH-mm-ss")}.txt";
            Console.WriteLine("Unzipping Files...");
            foreach (var file in newFiles)
            {
                if (Path.GetExtension(file) == ".zip" || Path.GetExtension(file) == ".ZIP")
                {
                    Decompress(source);
                    break;
                }
            }
            Console.WriteLine("Unzip Successful");
            newFiles = Directory.GetFiles(source);
            #endregion

            #region Decrypt
            int count = 0;
            if (decrypt == "on")
            {
                Console.WriteLine("Decrypting Files...");

                foreach (var file in newFiles)
                {
                    count++;
                    EncryptedChecker(file);
                    if (isEncrypted == true)
                        Decrypting(file);
                    isEncrypted = false;
                    double percentage = (double)count / newFiles.Length * 100;
                    if (count % 10 == 0)
                        Console.WriteLine($"Processing {percentage.ToString("0")}%");
                }
                Console.WriteLine("Decrypting Successful");
                newFiles = Directory.GetFiles(source);
            }
            #endregion

            if (newFiles.Length >= 1)
            {
                int countPassed = 0;
                int countFailed = 0;
                count = 0;
                bool match = false;
                StreamWriter sw;
                sw = File.CreateText(logFile);

                foreach (var fileOne in newFiles)
                {
                    count++;
                    foreach (var fileTwo in oldFiles)
                    {
                        //SameFileName
                        string mainFileOne = Path.GetFileName(fileOne);
                        string mainFileTwo = Path.GetFileName(fileTwo);

                        string newFileOne = Path.GetFileNameWithoutExtension(fileOne).Replace("PC", "").Replace("APPS", "");
                        string newFileTwo = Path.GetFileNameWithoutExtension(fileTwo).Replace("PC", "").Replace("APPS", "");

                        string subFileOne = Regex.Match(Path.GetFileNameWithoutExtension(newFileOne), "[a-zA-Z]+-[0-9]{1,10}[a-zA-Z]{1,7}").ToString(); //newFiles
                        string subFileTwo = Regex.Match(Path.GetFileNameWithoutExtension(newFileTwo), "[a-zA-Z]+-[0-9]{1,10}[a-zA-Z]{1,7}").ToString(); //oldFiles

                        bool isMatch = Regex.IsMatch(Path.GetFileNameWithoutExtension(fileTwo), "[a-zA-Z]+-[0-9]{1,10}[a-oq-zA-OQ-Za-oq-z]{0,7}");
                        //match (subFile = OPSH-###CD)
                        //newFiles =  OPSH-##########ABCDEEE - PC = "" - APPS = ""
                        //oldFiles = OPSH-##########ABCDEEE
                        if (isMatch && subFileOne == subFileTwo)
                        {
                            try
                            {
                                FileProcess(fileOne, fileTwo, failedPath);
                                break;
                            }
                            catch (Exception x)
                            {
                                sw.WriteLine($"{x.Message} {Path.GetFileName(fileOne)}");
                                Console.WriteLine($"\nError Message: {x.Message} - {Path.GetFileName(fileOne)}");
                            }
                        }
                        else if (!isMatch && mainFileOne == mainFileTwo) //!match(fileOne = string)
                        {
                            try
                            {
                                FileProcess(fileOne, fileTwo, failedPath);
                                break;
                            }
                            catch (Exception x)
                            {
                                sw.WriteLine($"{x.Message} {Path.GetFileName(fileOne)}");
                                Console.WriteLine($"\nError Message: {x.Message} - {Path.GetFileName(fileOne)}");
                            }
                        }
                    }
                    if (isPassed == true) // true
                    {
                        sw.WriteLine($"{Path.GetFileName(fileOne)} is Passed");
                        countPassed++;
                    }
                    else if (isPassed == false)// false
                    {
                        sw.WriteLine($"{Path.GetFileName(fileOne)} is Failed");
                        countFailed++;
                    }
                    Console.WriteLine($"{count}/{newFiles.Length} - {Path.GetFileName(fileOne)}");
                }
                sw.WriteLine($"Count Passed: {countPassed}\nCount Failed: {countFailed}");
                sw.Close();
            }
        }
        public void FileProcess(string fileOne, string fileTwo, string failedPath)
        {
            #region Process of Files
            if (string.Equals((Path.GetExtension(fileOne)), ".txt", StringComparison.OrdinalIgnoreCase))
            {
                CompareTxtFiles(fileOne, fileTwo, failedPath);
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xls", StringComparison.OrdinalIgnoreCase))
            {
                CompareXLSFiles(fileOne, fileTwo, failedPath);
                CloseExcel();
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                CompareXLSXFiles(fileOne, fileTwo, failedPath);
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".csv", StringComparison.OrdinalIgnoreCase))
            {
                CompareCSVFiles(fileOne, fileTwo, failedPath);
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                ComparePDFFiles(fileOne, fileTwo, failedPath);
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xml", StringComparison.OrdinalIgnoreCase))
            {
                CompareXMLFiles(fileOne, fileTwo, failedPath);
            }
            #endregion
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

        #region Unzip & Decrypt
        public void Decompress(string zip)
        {
            string temp = Path.Combine(source + @"Temp\");
            string[] files = Directory.GetFiles(zip, "*.zip", SearchOption.AllDirectories);
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
        #region FileName Compare
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
        #endregion

        #region Text Comparison
        public void CompareTxtFiles(string fileOne, string fileTwo, string failedPath)
        {
            string[] lines = File.ReadAllLines(fileOne);
            string[] lines2 = File.ReadAllLines(fileTwo);

            List<string> newFiles = new List<string>();
            List<string> oldFiles = new List<string>();
            foreach (string line in lines)
            {
                newFiles.Add(line);
            }
            foreach (var line2 in lines2)
            {
                oldFiles.Add(line2);
            }
            if (Enumerable.SequenceEqual(newFiles, oldFiles) == true)
                isPassed = true;
            else
            {
                File.Copy(fileOne, failedPath + Path.GetFileName("/" + fileOne));
                isPassed = false;
            }
        }
        #endregion

        #region Excel Comparison
        int rowCount = 0;
        int colCount = 0;
        public void CompareXLSFiles(string fileOne, string fileTwo, string failedPath)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileOne);

            Excel.Application xlAppTwo = new Excel.Application();
            Excel.Workbook xlWorkbookTwo = xlAppTwo.Workbooks.Open(fileTwo);

            Excel._Worksheet xlWorksheet;
            Excel._Worksheet xlWorksheetTwo;

            bool worksheetOne = xlWorkbook.Sheets.Count > 1;
            bool worksheetTwo = xlWorkbookTwo.Sheets.Count > 1;
            if (worksheetOne && worksheetTwo)
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

            for (int i = 1; i <= row; i++) // row {1}
            {
                for (int j = 1; j <= col; j++) // col {A}
                {
                    if (!String.IsNullOrEmpty(xlWorksheet.Cells[i, j].Text.ToString()) || !String.IsNullOrWhiteSpace(xlWorksheet.Cells[i, j].Text.ToString()))
                    {
                        var wsOne = xlWorksheet.Cells[i, j].Value;
                        string wsOnes = wsOne.ToString();
                        if (Regex.IsMatch(wsOnes, @"^-?[0-9]\d*(\.\d+)?$"))
                        {
                            double one = double.Parse($"{xlWorksheet.Cells[i, j].Value}");
                            double two = double.Parse($"{xlWorksheetTwo.Cells[i, j].Value}");
                            if (one != two)
                            {
                                string getFileName = Path.GetFileName(fileOne);
                                File.Copy(fileOne, failedPath + Path.GetFileName(fileOne));
                                cancel = true;
                                break;
                            }
                        }
                        else
                        {
                            if (xlWorksheet.Cells[i, j].Text.ToString() != xlWorksheetTwo.Cells[i, j].Text.ToString())
                            {
                                string getFileName = Path.GetFileName(fileOne);
                                File.Copy(fileOne, failedPath + Path.GetFileName(fileOne));
                                cancel = true;
                                break;
                            }
                        }
                    }
                }
                if (cancel)
                {
                    isPassed = false;
                    break;
                }
                else
                    isPassed = true;
            }
            xlApp.Workbooks.Close();
            xlApp.Quit();
            xlAppTwo.Workbooks.Close();
            xlAppTwo.Quit();
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlWorkbookTwo);
        }
        #endregion

        #region XLSX Comparison
        public void CompareXLSXFiles(string fileOne, string fileTwo, string failedPath)
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
                    if (wsOne.Cells[i, j].Text != wsTwo.Cells[i, j].Text)
                    {
                        //using (ExcelRange rng = wsOne.Cells[i, j])
                        //{
                        //    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                        //}
                        isFailed = true;
                    }
                }
            }
            excelOne.Save();
            if (isFailed)
            {
                File.Copy(fileOne, failedPath + Path.GetFileName(fileTwo));
                isPassed = false;
            }
            else
                isPassed = true;
        }
        #endregion

        #region CSV Comparison
        public void CompareCSVFiles(string fileOne, string fileTwo, string failedPath)
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
                        isPassed = false;
                        break;
                    }
                    else
                        isPassed = true;
                }
            }
        }
        #endregion

        #region PDF Comparison
        public void ComparePDFFiles(string fileOne, string fileTwo, string failedPath)
        {
            iText.PdfReader pdfOne = new iText.PdfReader(fileOne);
            iText.PdfReader pdfTwo = new iText.PdfReader(fileTwo);
            int pageFileOne = pdfOne.NumberOfPages;
            string line;
            string dataOne;
            string dataTwo;
            for (int i = 1; i <= pageFileOne; i++)
            {
                if (i <= 1)
                {
                    dataOne = iTextParser.PdfTextExtractor.GetTextFromPage(pdfOne, i, new iTextParser.LocationTextExtractionStrategy());
                    dataTwo = iTextParser.PdfTextExtractor.GetTextFromPage(pdfTwo, i, new iTextParser.LocationTextExtractionStrategy());
                    dataOne = Regex.Replace(dataOne, @"(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+\s-\s(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+", "");//remove date period
                    dataTwo = Regex.Replace(dataTwo, @"(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+\s-\s(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+", "");//remove page# 
                    dataTwo = Regex.Replace(dataTwo, @"[A-Za-z]+\s[A-Za-z]+\s[A-Za-z]+:\s[0-9]{2}:[0-9]{2}:[0-9]{2}(\.[0-9]{1,3})?", "");

                    if (Enumerable.SequenceEqual(dataOne, dataTwo) == true)
                        isPassed = true;
                    else
                    {
                        File.Copy(fileOne, failedPath + Path.GetFileName("/" + fileOne));
                        isPassed = false;
                        break;
                    }
                }
            }
        }
        #endregion

        #region XML Comparison
        public void CompareXMLFiles(string fileOne, string fileTwo, string failedPath)
        {
            string[] fileOneLines = File.ReadAllLines(fileOne);
            string[] fileTwoLines = File.ReadAllLines(fileTwo);

            List<string> newFiles = new List<string>();
            List<string> oldFiles = new List<string>();

            foreach (var line in fileOneLines)
            {
                newFiles.Add(line);
            }
            foreach (var line in fileTwoLines)
            {
                oldFiles.Add(line);
            }
            if (Enumerable.SequenceEqual(newFiles, oldFiles) == true)
                isPassed = true;
            else
            {
                File.Copy(fileOne, failedPath + Path.GetFileName("/" + fileOne));
                isPassed = false;
            }
        }
        #endregion
    }
}
