using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using iText = iTextSharp.text.pdf;
using iTextParser = iTextSharp.text.pdf.parser;

namespace AT_ER_OutputFiles
{
    internal class AT_ER_OutputFile
    {
        #region Declarations
        FileTool fileTool = new FileTool();
        string source = ConfigurationSettings.AppSettings["SourcePath"]; 
        string temp = ConfigurationSettings.AppSettings["Temp"];
        string e2eProcess = ConfigurationSettings.AppSettings["E2EProcess"].ToLower();
        bool isPassed = false;
        #endregion
        
        #region Process of Files
        public void ProcessOfFiles()
        {
            string destination = ConfigurationSettings.AppSettings["DestinationPath"];
            string failedPath = ConfigurationSettings.AppSettings["FailedPath"];
            
            string[] newFiles = Directory.GetFiles(source);
            string[] oldFiles = Directory.GetFiles(destination);

            if (e2eProcess == "true")
                fileTool.EndToEndProcess();

            newFiles = Directory.GetFiles(source);
            #region Decompress
            string logFile = $@"{temp}LOG-{DateTime.Now.ToString("MM-d-yy-HH-mm-ss")}.txt";
            Console.WriteLine("Unzipping Files...");
            foreach (var file in newFiles)
            {
                if (Path.GetExtension(file) == ".zip" || Path.GetExtension(file) == ".ZIP")
                {
                    fileTool.Decompress(source);
                    break;
                }
            }
            Console.WriteLine("Unzip Successful");
            newFiles = Directory.GetFiles(source);
            #endregion
            #region Decrypt
            int count = 0;
            Console.WriteLine("Decrypting Files...");

            foreach (var file in newFiles)
            {
                count++;
                fileTool.EncryptedChecker(file);
                if (fileTool.isEncrypted == true)
                    fileTool.Decrypting(file);
                fileTool.isEncrypted = false;
                double percentage = (double)count / newFiles.Length * 100;
                if (count % 10 == 0)
                    Console.WriteLine($"Processing {percentage.ToString("0")}%");
            }
            Console.WriteLine("Decrypting Successful");
            newFiles = Directory.GetFiles(source);
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
                        string mainFileOne = Path.GetFileName(fileOne); //2
                        string mainFileTwo = Path.GetFileName(fileTwo); //2

                        string newFileOne = Path.GetFileNameWithoutExtension(fileOne).Replace("PC", "").Replace("APPS", "");
                        string newFileTwo = Path.GetFileNameWithoutExtension(fileTwo).Replace("PC", "").Replace("APPS", "");

                        string subFileOne = Regex.Match(Path.GetFileNameWithoutExtension(newFileOne), "[a-zA-Z]+-[0-9]{1,10}[a-zA-Z]{1,7}").ToString(); //newFiles 1
                        string subFileTwo = Regex.Match(Path.GetFileNameWithoutExtension(newFileTwo), "[a-zA-Z]+-[0-9]{1,10}[a-zA-Z]{1,7}").ToString(); //oldFiles 1

                        bool isMatch = Regex.IsMatch(Path.GetFileNameWithoutExtension(fileTwo), "[a-zA-Z]+-[0-9]{1,10}[a-oq-zA-OQ-Za-oq-z]{0,7}");

                        string removeDateOne = Regex.Replace(mainFileOne, @"\d", "");
                        string removeDateTwo = Regex.Replace(mainFileTwo, @"\d", "");

                        //match (subFile = OPSH-###CD)
                        //newFiles =  OPSH-##########ABCDEEE - PC = "" - APPS = ""
                        //oldFiles = OPSH-##########ABCDEEE
                        if (isMatch && subFileOne == subFileTwo) //1
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
                        //Start JM - 04/23/24
                        else if (!isMatch && removeDateOne == removeDateTwo)
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
                        //End JM - 04/23/24
                    }
                    if (isPassed == true) // true
                    {
                        sw.WriteLine($"{Path.GetFileName(fileOne)} | Result: Passed");
                        countPassed++;
                    }
                    else if (isPassed == false)// false
                    {
                        sw.WriteLine($"{Path.GetFileName(fileOne)} | Result: Failed");
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
            if (string.Equals((Path.GetExtension(fileOne)), ".txt", StringComparison.OrdinalIgnoreCase))
                CompareTxtFiles(fileOne, fileTwo, failedPath);
            else if (string.Equals((Path.GetExtension(fileOne)), ".xls", StringComparison.OrdinalIgnoreCase))
            {
                if (fileOne.Contains("EXFMT"))
                {
                    CompareXLSFiles(fileOne, fileTwo, failedPath);
                    fileTool.CloseExcel();
                }
                else
                    NewCompareXLSFiles(fileOne, fileTwo, failedPath);
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xlsx", StringComparison.OrdinalIgnoreCase))
                CompareXLSXFiles(fileOne, fileTwo, failedPath);
            else if (string.Equals((Path.GetExtension(fileOne)), ".csv", StringComparison.OrdinalIgnoreCase))
                CompareCSVFiles(fileOne, fileTwo, failedPath);
            else if (string.Equals((Path.GetExtension(fileOne)), ".pdf", StringComparison.OrdinalIgnoreCase))
                ComparePDFFiles(fileOne, fileTwo, failedPath);
            else if (string.Equals((Path.GetExtension(fileOne)), ".xml", StringComparison.OrdinalIgnoreCase))
                CompareXMLFiles(fileOne, fileTwo, failedPath);
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
        public void NewCompareXLSFiles(string fileOne, string fileTwo, string failedPath)
        {
            int rowCount = 0;
            int colCount = 0;
            bool cancel = false;

            string[] fileOneLines = File.ReadAllLines(fileOne);
            string[] fileTwoLines = File.ReadAllLines(fileTwo);

            int row = fileOneLines.Length;
            int col = fileOneLines[0].Split(',').Length;

            for (int i = 0; i < row; i++) // row index starts from 0 for arrays
            {
                string[] fileOneCols = fileOneLines[i].Split('\t');
                string[] fileTwoCols = fileTwoLines[i].Split('\t');

                for (int j = 0; j < col; j++) // col index starts from 0 for arrays
                {
                    if (!string.IsNullOrEmpty(fileOneCols[j]) || !string.IsNullOrWhiteSpace(fileOneCols[j]))
                    {
                        string wsOnes = fileOneCols[j];
                        if (Regex.IsMatch(wsOnes, @"^-?[0-9]\d*(\.\d+)?$"))
                        {
                            double one = double.Parse(fileOneCols[j]);
                            double two = double.Parse(fileTwoCols[j]);
                            if (one != two)
                            {
                                string getFileName = Path.GetFileName(fileOne);
                                File.Copy(fileOne, Path.Combine(failedPath, Path.GetFileName(fileOne)), true);
                                cancel = true;
                                break;
                            }
                        }
                        else
                        {
                            if (fileOneCols[j] != fileTwoCols[j])
                            {
                                string getFileName = Path.GetFileName(fileOne);
                                File.Copy(fileOne, Path.Combine(failedPath, Path.GetFileName(fileOne)), true);
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
                {
                    isPassed = true;
                }
            }

        }
        public void CompareXLSFiles(string fileOne, string fileTwo, string failedPath)
        {
            int rowCount = 0;
            int colCount = 0;
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
            var wsOne = excelOne.Workbook.Worksheets[0]; //new
            var wsTwo = excelTwo.Workbook.Worksheets[0]; //old
            int row = wsOne.Dimension.End.Row;
            int col = wsOne.Dimension.End.Column;
            bool isFailed = false;
            for (int i = 1; i <= row; i++)
            {
                for (int j = 1; j <= col; j++)
                {
                    if (wsOne.Cells[i, j].Text != wsTwo.Cells[i, j].Text)
                    {
                        isFailed = true;
                        break;
                    }
                }
                if (isFailed == true)
                    break;
            }
            if (isFailed)
            {
                string test = failedPath + Path.GetFileName(fileTwo);
                File.Copy(fileOne, failedPath + Path.GetFileName(fileOne));
                isPassed = false;
                return;
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
                    //remove date period
                    dataOne = Regex.Replace(dataOne, @"(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+\s-\s(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+", "");//remove date period
                    dataTwo = Regex.Replace(dataTwo, @"(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+\s-\s(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+", "");//remove date period

                    //dataOne = Regex.Replace(dataOne, @"(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+\s-\s(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+", "");//remove page# 
                    //dataTwo = Regex.Replace(dataTwo, @"(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+\s-\s(0?[1-9]|[12][0-9]|3[01])\.(0?[1-9]|[1][0-2])\.[0-9]+", "");//remove page# 

                    //remove execute time 
                    dataOne = Regex.Replace(dataOne, @"[A-Za-z]+\s[A-Za-z]+\s[A-Za-z]+:\s[0-9]{2}:[0-9]{2}:[0-9]{2}(\.[0-9]{1,3})?", "");
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
