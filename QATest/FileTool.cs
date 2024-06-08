using OfficeOpenXml;
using Renci.SshNet;
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
        string processOne = ConfigurationSettings.AppSettings["SourcePath"];
        string processTwo = ConfigurationSettings.AppSettings["EncodingPath"];
        string temp = ConfigurationSettings.AppSettings["Temp"];
        string server = ConfigurationSettings.AppSettings["Server"];
        string e2eProcess = ConfigurationSettings.AppSettings["E2EProcess"].ToLower();

        public bool isEncrypted = false;

        public void EndToEndProcess()
        {
            ExecuteScript executeScript = new ExecuteScript();
            if (e2eProcess == "true")
            {
                #region Execute Reports
                executeScript.LoginSapGui();
                executeScript.RunSapScripting();
                #endregion
                #region Timer mins
                Timer(3);
                #endregion
                #region GMAIL
                executeScript.DownloadGmail();
                #endregion
                #region SFTP
                SFTPConnect();
                #endregion
                #region Backup File
                FileBackup();
                #endregion
            }
        }

        #region Comparison/Encoding Only
        public bool DirectProcess()
        {
            string[] fileCount;
            if (pType == "1")
            {
                fileCount = Directory.GetFiles(processOne);
                if (fileCount.Length > 0)
                    return true;
            }
            else if (pType == "2")
            {
                fileCount = Directory.GetFiles(processTwo);
                if (fileCount.Length > 0)
                    return true;
            }
            return false;
        }
        #endregion

        #region System Identify
        public string SystemIdentify()
        {
            if (server == "52.74.184.114")
                return "MT4";
            else if (server == "52.74.10.149")
                return "MT5";
            else if (server == "52.221.171.219")
                return "MT6";
            else
                return null;
        }
        #endregion

        #region File Backup Comparison or Encoding
        public void FileBackup()
        {
            Console.WriteLine("File Backup");
            DateTime currentDate = DateTime.Now;
            string formattedDate = currentDate.ToString("MMddyyyy");
            string fullPath = null;
            string fileName = null;
            string system = SystemIdentify();
            string[] newFilesOne = Directory.GetFiles(processOne);
            string[] newFilesTwo = Directory.GetFiles(processTwo);
            if (pType == "1")
            {
                foreach (var file in newFilesOne)
                {
                    fullPath = Path.Combine(processOne, formattedDate + $"-{system}");
                    if (!Directory.Exists(fullPath))
                        Directory.CreateDirectory(fullPath);
                    fileName = Path.GetFileName(file);
                    File.Copy(file, Path.Combine(fullPath, fileName));
                }
            }

            else if (pType == "2")
                foreach (var file in newFilesTwo)
                {
                    fullPath = Path.Combine(processTwo, formattedDate + $"-{system}");
                    if (!Directory.Exists(fullPath))
                        Directory.CreateDirectory(fullPath);
                    fileName = Path.GetFileName(file);
                    File.Copy(file, Path.Combine(fullPath, fileName));
                }
            Console.WriteLine("File Backup Done");
        }
        #endregion

        #region Path creator
        public void PathCreator()
        {
            if (!Directory.Exists(temp))
                Directory.CreateDirectory(temp);
            string oldFiles = Path.Combine(temp, "oldfiles");
            string newFiles = Path.Combine(temp, "newfiles");
            string failedPath = Path.Combine(temp, "failedpath");
            if (!Directory.Exists(oldFiles))
                Directory.CreateDirectory(oldFiles);
            if (!Directory.Exists(newFiles))
                Directory.CreateDirectory(newFiles);
            if (!Directory.Exists(failedPath))
                Directory.CreateDirectory(failedPath);
        }
        #endregion

        #region CloseExcel
        public void CloseExcel(Excel.Application ExcelApplication = null)
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

        #region Process Type Changer
        public void ProcessTypeChanger()
        {
            if (pType == "2")
                source = ConfigurationSettings.AppSettings["EncodingPath"];
        }
        #endregion

        #region Timer
        public void Timer(int mins)
        {
            int remainingTimeInSeconds = mins * 60; // minutes in seconds

            Console.WriteLine($"Timer started for {mins} minutes.");

            while (remainingTimeInSeconds > 0)
            {
                // Calculate minutes and seconds
                int minutes = remainingTimeInSeconds / 60;
                int seconds = remainingTimeInSeconds % 60;

                // Display remaining time
                Console.WriteLine($"Time remaining: {minutes:00}:{seconds:00}");

                // Wait for one second
                System.Threading.Thread.Sleep(1000);

                // Decrement remaining time
                remainingTimeInSeconds--;
            }

            Console.WriteLine("Timer expired!");
        }
        #endregion

        #region SFTP Connect
        public void SFTPConnect()
        {
            Console.WriteLine("\nStarting Moving of Files from Server to Local");
            string destinationPath = temp;
            string pType = ConfigurationSettings.AppSettings["ProcessType"];
            string comparisonPath = ConfigurationSettings.AppSettings["SourcePath"];
            string encodingPath = ConfigurationSettings.AppSettings["EncodingPath"];

            if (pType == "1")
                destinationPath = comparisonPath;
            else if (pType == "2")
                destinationPath = encodingPath;
            else if (pType == "5")
                destinationPath = temp;

            string sftpPath = ConfigurationSettings.AppSettings["SFTP"];

            string username = ConfigurationSettings.AppSettings["Username"];
            string password = ConfigurationSettings.AppSettings["Password"];
            int port = int.Parse(ConfigurationSettings.AppSettings["Port"]);
            int count = 0;

            List<string> fileList = new List<string>();

            using (var client = new SftpClient(server, port, username, password))
            {
                client.Connect();

                var files = client.ListDirectory(sftpPath);
                foreach (var file in files)
                {
                    // Skip directories
                    if (file.IsDirectory)
                        continue;
                    // Construct remote file path
                    string remoteFilePath = sftpPath + "/" + file.Name;
                    // Construct local destination file path
                    string localDestinationPath = Path.Combine(destinationPath, file.Name);
                    // Download file from SFTP to local directory
                    using (var fileStream = File.Create(localDestinationPath))
                    {
                        client.DownloadFile(remoteFilePath, fileStream);
                        client.DeleteFile(remoteFilePath);
                    }
                    count++;
                }
                client.Disconnect();
            }
            Console.WriteLine($"\nCompleted Moving {count} files from Server");
        }
        #endregion

        #region Decompress
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
        #endregion

        #region Checking if Encrypted or Not
        public void EncryptedChecker(string fileOne)
        {
            isEncrypted = false;
            if (string.Equals(Path.GetExtension(fileOne), ".txt", StringComparison.OrdinalIgnoreCase) || string.Equals(Path.GetExtension(fileOne), ".xls", StringComparison.OrdinalIgnoreCase))
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
            //else if (string.Equals(Path.GetExtension(fileOne), ".xls", StringComparison.OrdinalIgnoreCase))
            //{
            //    Excel.Application xlApp = new Excel.Application();
            //    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileOne);
            //    Excel.Application xlAppTwo = new Excel.Application();
            //    Excel._Worksheet xlWorksheet;
            //    xlWorksheet = xlWorkbook.Sheets[1];

            //    for (int i = 1; i <= xlWorksheet.UsedRange.Rows.Count; i++) // row {1}
            //    {
            //        for (int j = 1; j <= xlWorksheet.UsedRange.Columns.Count; j++) // col {A}
            //        {
            //            if (!string.IsNullOrEmpty(xlWorksheet.Cells[i, j].Text.ToString()) || !string.IsNullOrWhiteSpace(xlWorksheet.Cells[i, j].Text.ToString())) { }
            //            if (xlWorksheet.Cells[i, j].Text.ToString().Contains("AES"))
            //            {
            //                isEncrypted = true;
            //                break;
            //            }
            //        }
            //    }
            //    xlApp.Workbooks.Close();
            //    xlApp.Quit();
            //    CloseExcel();
            //}
            else if (string.Equals(Path.GetExtension(fileOne), ".xlsx", StringComparison.OrdinalIgnoreCase))
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
            else if (string.Equals(Path.GetExtension(fileOne), ".csv", StringComparison.OrdinalIgnoreCase))
            {
                using (StreamReader f1 = new StreamReader(fileOne))
                {
                    var line1 = f1.ReadLine();
                    if (line1.Contains("AES"))
                        isEncrypted = true;
                }
            }
            else if (string.Equals(Path.GetExtension(fileOne), ".pdf", StringComparison.OrdinalIgnoreCase))
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
            else if (string.Equals(Path.GetExtension(fileOne), ".xml", StringComparison.OrdinalIgnoreCase))
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
        #endregion

        #region Decrypting
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
    }
}
