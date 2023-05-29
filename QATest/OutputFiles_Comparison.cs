using OfficeOpenXml;
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
            if(!Directory.Exists(destination))
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
            //PathCreator();
            string destination = ConfigurationSettings.AppSettings["DestinationPath"]; //old

            string[] newFiles = Directory.GetFiles(source);
            string[] oldFiles = Directory.GetFiles(destination);

            string logFile = $@"c:\temp\LOG-{DateTime.Now.ToString("MM-d-yy-HH-mm-ss")}.txt";
            foreach (var file in newFiles)
            {
                if (Path.GetExtension(file) == ".zip" || Path.GetExtension(file) == ".ZIP")
                {
                    Decompress(source);
                    break;
                }
            }
            newFiles = Directory.GetFiles(source);

            string passedPath = ConfigurationSettings.AppSettings["PassedPath"];
            string failedPath = ConfigurationSettings.AppSettings["FailedPath"];

            int countPassed = 0;
            int countFailed = 0;
            int count = 0;
            bool match = false;

            if (newFiles.Length >= 1)
            {
                StreamWriter sw;
                sw = File.CreateText(logFile);
                string newName = null;

                foreach (var fileOne in newFiles)
                {
                    count++;
                    foreach (var fileTwo in oldFiles)
                    {
                        //SameFileName
                        string mainFileOne = Path.GetFileName(fileOne);
                        string mainFileTwo = Path.GetFileName(fileTwo);

                        string subFileOne = Regex.Match(Path.GetFileNameWithoutExtension(fileOne), "[a-zA-Z]+-[0-9]{1,10}[a-zA-Z]{1,7}").ToString().Replace("PC", "").Replace("APPS", "");
                        string subFileTwo = Regex.Match(Path.GetFileNameWithoutExtension(fileTwo), "[a-zA-Z]+-[0-9]{1,10}[a-zA-Z]{1,7}").ToString();

                        bool isMatch = Regex.IsMatch(Path.GetFileNameWithoutExtension(fileTwo), "[a-zA-Z]+-[0-9]{1,10}[a-zA-Z]{1,7}");



                        if (isMatch && subFileOne == subFileTwo)  //match (subFile = OPSH-###CD)
                        {
                            try
                            {
                                newName = FileProcess(fileOne, fileTwo, passedPath, failedPath);
                                match = true;
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
                                newName = FileProcess(fileOne, fileTwo, passedPath, failedPath);
                                match = true;
                                break;
                            }
                            catch (Exception x)
                            {
                                sw.WriteLine($"{x.Message} {Path.GetFileName(fileOne)}");
                                Console.WriteLine($"\nError Message: {x.Message} - {Path.GetFileName(fileOne)}");
                            }
                        }
                        else
                            match = false;
                    }
                    

                    if (isPassed == true && match == true) // true
                    {
                        
                        if (match == true)
                            Console.WriteLine($"{count}/{newFiles.Length} - {Path.GetFileName(newName)} - Passed");
                        sw.WriteLine($"{newName} is Passed");
                        countPassed++;
                    }
                    else if (isPassed == false) // false
                    {
                        if (match == true)
                        {
                            Console.WriteLine($"{count}/{newFiles.Length} - {Path.GetFileName(newName)} - Failed");
                            sw.WriteLine($"{newName} is Failed");
                        }
                        else if (match == false)
                        {
                            Console.WriteLine($"{count}/{newFiles.Length} - {Path.GetFileName(fileOne)} does not exist in oldFiles");
                            sw.WriteLine($"{fileOne} is Failed");
                        }
                        countFailed++;
                    }
                    match = false; 
                }
                sw.WriteLine($"Count Passed: {countPassed}\nCount Failed: {countFailed}");
                sw.Close();
            }
        }
        public string FileProcess(string fileOne, string fileTwo, string passedPath, string failedPath)
        {
            #region Process of Files
            if (string.Equals((Path.GetExtension(fileOne)), ".txt", StringComparison.OrdinalIgnoreCase))
            {
                EncryptedChecker(fileOne);
                if (isEncrypted == true)
                    fileOne = Decrypting(fileOne);
                CompareTxtFiles(fileOne, fileTwo, passedPath, failedPath);
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xls", StringComparison.OrdinalIgnoreCase))
            {
                EncryptedChecker(fileOne);
                if (isEncrypted == true)
                    fileOne = Decrypting(fileOne);
                CloseExcel();
                CompareXLSFiles(fileOne, fileTwo, passedPath, failedPath);
                CloseExcel();
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                EncryptedChecker(fileOne);
                if (isEncrypted == true)
                    fileOne = Decrypting(fileOne);
                CompareXLSXFiles(fileOne, fileTwo, passedPath, failedPath);
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".csv", StringComparison.OrdinalIgnoreCase))
            {
                EncryptedChecker(fileOne);
                if (isEncrypted == true)
                    fileOne = Decrypting(fileOne);
                CompareCSVFiles(fileOne, fileTwo, passedPath, failedPath);
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                EncryptedChecker(fileOne);
                if (isEncrypted == true)
                    fileOne = Decrypting(fileOne);
                ComparePDFFiles(fileOne, fileTwo, passedPath, failedPath);
            }
            else if (string.Equals((Path.GetExtension(fileOne)), ".xml", StringComparison.OrdinalIgnoreCase))
            {
                EncryptedChecker(fileOne);
                if(isEncrypted == true)
                    fileOne = Decrypting(fileOne);
                CompareXMLFiles(fileOne, fileTwo, passedPath, failedPath);
            }
            source = ConfigurationSettings.AppSettings["SourcePath"];
            isEncrypted = false;
            return fileOne;
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
                string newPathFile = null;
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
                    newPathFile = source + withoutExt + itsExt;
                    if (!File.Exists(newPathFile))
                    {
                        File.Move(fileTemp, newPathFile);
                    }
                    else if (File.Exists(newPathFile))
                    {
                        newPathFile = path + "\\" + withoutExt + $"({fileIncrement}){itsExt}";
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
        public void CompareTxtFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
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
        public void CompareXLSFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
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
                        if (xlWorksheet.Cells[i, j].Text.ToString().Contains("AES"))
                            throw new Exception("lalala");

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
        public void ComparePDFFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
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
                    if (Enumerable.SequenceEqual(dataOne, dataTwo) == true)
                        isPassed = true;
                    else
                    {
                        File.Copy(fileOne, failedPath + Path.GetFileName("/" + fileOne));
                        isPassed = false;
                    }
                }
            }
        }
        #endregion

        #region XML Comparison
        public void CompareXMLFiles(string fileOne, string fileTwo, string passedPath, string failedPath)
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
            if(Enumerable.SequenceEqual(newFiles, oldFiles) == true)
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
