using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QATest
{
    internal class StartProcess
    {
        public void FileAutomation()
        {
            string processType = ConfigurationSettings.AppSettings["ProcessType"];

            string source = ConfigurationSettings.AppSettings["SourcePath"];
            string destination = ConfigurationSettings.AppSettings["DestinationPath"];

            string[] newFiles = Directory.GetFiles(source);
            string[] oldFiles = Directory.GetFiles(destination);

            string passedPath = ConfigurationSettings.AppSettings["PassedPath"];
            string failedPath = ConfigurationSettings.AppSettings["FailedPath"];

            if (processType == "1") //ProcessType - 1: TextFile
            {
                CompareTxtFiles(newFiles, oldFiles, passedPath, failedPath);
            }
            else if (processType == "2") //ProcessType - 2: Compress and Text
            {
                Decompress(source);
                newFiles = Directory.GetFiles(source);
                CompareTxtFiles(newFiles, oldFiles, passedPath, failedPath);
            }
            else if (processType == "3") //ProcessType - 3: Encrypted and Text
            {
                Decrypting(source);
                CompareTxtFiles(newFiles, oldFiles, passedPath, failedPath);
            }
            else if (processType == "4") //ProcessType - 4: Compress, Encrypted and Text
            {
                Decompress(source);
                Decrypting(source);
                newFiles = Directory.GetFiles(source);
                CompareTxtFiles(newFiles, oldFiles, passedPath, failedPath);
            }
            else if (processType == "5") //ProcessType - 5: Filename Comparison
            {
                CompareFileName(newFiles, oldFiles, passedPath, failedPath);
            }
            else if (processType == "6") //ProcessType - 6: Excel Files Comparison
            {
                CompareExcelFiles(newFiles, oldFiles, passedPath, failedPath);
            }
            else if (processType == "7")
            {
                CompareCSVFiles(newFiles, oldFiles, passedPath, failedPath);
            }
        }
        public void CompareTxtFiles(string[] newFiles, string[] oldFiles, string passedPath, string failedPath)
        {

            foreach (var oneFile in oldFiles)
            {
                foreach (var twoFile in newFiles)
                {
                    if (Path.GetFileName(oneFile) == Path.GetFileName(twoFile))
                    {
                        string[] lines = File.ReadAllLines(oneFile);
                        string[] lines2 = File.ReadAllLines(twoFile);

                        List<string> fileOne = new List<string>();
                        List<string> fileTwo = new List<string>();
                        foreach (string line in lines)
                        {
                            fileOne.Add(line);
                        }
                        foreach (var line2 in lines2)
                        {
                            fileTwo.Add(line2);
                        }
                        if (Enumerable.SequenceEqual(fileOne, fileTwo) == true)
                        {
                            File.Copy(oneFile, passedPath + Path.GetFileName(oneFile));
                        }
                        else
                        {
                            File.Copy(oneFile, failedPath + Path.GetFileName("/" + oneFile));
                        }
                        break;
                    }
                }
            }
        }
        public void Decompress(string source)
        {
            string temp = Path.Combine(source + @"Temp\");
            string[] files = Directory.GetFiles(source);
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
        public void Decrypting(string source)
        {
            string[] files = Directory.GetFiles(source, "*txt");
            foreach (var file in files)
            {
                File.Copy(file, Path.ChangeExtension(file, ".aes"));
                string fileName = Path.GetFileNameWithoutExtension(file);
                using (var outputfile = System.IO.File.OpenWrite(file))
                {
                    using (var inputfile = System.IO.File.OpenRead(source + fileName + ".aes"))
                    using (var encStream = new SharpAESCrypt.SharpAESCrypt("1", inputfile, SharpAESCrypt.OperationMode.Decrypt))
                    {
                        encStream.CopyTo(outputfile);
                    }
                }
                string remExt = file.Remove(file.Length - 4);
                File.Delete(remExt + ".aes");
            }
            string[] aesFiles = Directory.GetFiles(source, "*txt");
            foreach (var file in aesFiles)
            {
                string[] lines = File.ReadAllLines(file);
                List<string> fixedLines = new List<string>();
                int countLine = 0;
                foreach (var line in lines)
                {
                    countLine++;
                    if (countLine <= 9)
                    {
                        fixedLines.Add(line);
                    }
                }
                File.Delete(file);
                System.IO.File.WriteAllLines(file, fixedLines);
            }
        }
        public void CompareFileName(string[] newFiles, string[] oldFiles, string passedPath, string failedPath)
        {
            int oneCounter = 0;

            foreach (var oneFile in oldFiles)
            {
                oneCounter++;
                int twoCounter = 1;


                foreach (var twoFile in newFiles)
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
        public void CompareExcelFiles(string[] newFiles, string[] oldFiles, string passedPath, string failedPath)
        {
            foreach (var fileOne in newFiles)
            {
                foreach (var fileTwo in oldFiles)
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
                                if (wsOne.Cells[i, j].Text == wsTwo.Cells[i, j].Text)
                                {

                                }
                                if (wsOne.Cells[i, j].Text != wsTwo.Cells[i, j].Text)
                                {
                                    using (ExcelRange rng = wsOne.Cells[i, j])
                                    {
                                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        rng.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                                    }
                                    isFailed = true;
                                }
                            }
                        }
                        excelOne.Save();
                        if (isFailed)
                        {
                            File.Move(fileOne, failedPath + Path.GetFileName(fileOne));
                        }
                        break;
                    }
                }
            }
        }
        public void CompareCSVFiles(string[] newFiles, string[] oldFiles, string passedPath, string failedPath)
        {
            foreach (var fileOne in newFiles)
            {
                foreach (var fileTwo in oldFiles)
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
        public void DecryptingExcelFiles(string source)
        {
            var files = Directory.GetFiles(source, "*.*", SearchOption.AllDirectories)
            .Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsx") || s.EndsWith(".csv"));
            string extension = null;
            foreach (var file in files)
            {
                if (Path.GetExtension(file) == ".xls")
                    extension = ".xls";
                if (Path.GetExtension(file) == ".xlsx")
                    extension = ".xlsx";
                if (Path.GetExtension(file) == ".csv")
                    extension = ".csv";

                File.Copy(file, Path.ChangeExtension(file, ".aes"));
                string fileName = Path.GetFileNameWithoutExtension(file);
                var temp = Directory.CreateDirectory(source + @"temp\");
                using (var outputfile = System.IO.File.OpenWrite(temp + fileName + extension))
                {
                    using (var inputfile = System.IO.File.OpenRead(source + fileName + ".aes"))
                    using (var encStream = new SharpAESCrypt.SharpAESCrypt("1", inputfile, SharpAESCrypt.OperationMode.Decrypt))
                    {
                        encStream.CopyTo(outputfile);
                    }
                }
                File.Delete(file);
                File.Delete(source + fileName + ".aes");
                File.Copy(temp + fileName + extension, source + fileName + extension);
                Directory.Delete(source + @"temp\", true);
            }
        }
    }
}