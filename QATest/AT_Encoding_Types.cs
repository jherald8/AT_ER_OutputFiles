using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AT_ER_OutputFiles
{
    internal class AT_Encoding_Types
    {
        public void Process()
        {
            FileTool fileTool = new FileTool();
            string temp = ConfigurationSettings.AppSettings["Temp"];
            string source = ConfigurationSettings.AppSettings["EncodingPath"];
            string logFile = $@"{temp}EncodingLog{DateTime.Now.ToString("yyyyMdHHmmss")}.txt";

            #region GMAIL
            ExecutePython executePython = new ExecutePython();
            executePython.ExecutingPython();
            #endregion
            #region SFTP
            fileTool.SFTPConnect();
            #endregion

            ListOfEncoding();
            StreamWriter sw;
            sw = File.CreateText(logFile);
            Console.WriteLine("Unzipping Files...");
            foreach (string file in Directory.GetFiles(source))
                if (Path.GetExtension(file) == ".zip" || Path.GetExtension(file) == ".ZIP")
                {
                    fileTool.Decompress(source);
                    break;
                }
            Console.WriteLine("Unzip Successful");

            int count = 0;
            bool isEncrypted = false;
            Console.WriteLine("Decrypting Files...");
            foreach (var file in Directory.GetFiles(source))
            {
                count++;
                fileTool.EncryptedChecker(file);
                if (fileTool.isEncrypted == true)
                    fileTool.Decrypting(file);
                isEncrypted = false;
                double percentage = (double)count / Directory.GetFiles(source).Length * 100;
                if (count % 10 == 0)
                    Console.WriteLine($"Processing {percentage.ToString("0")}%");
            }
            Console.WriteLine("Decrypting Successful");

            fileTool.FileBackup();

            foreach (var file in Directory.GetFiles(source))
                DetectFileEncoding(file);
            MatchList(sw);
        }
        public Encoding DetectFileEncoding(string filePath)
        {
            FilesEncoding filesEncoding = new FilesEncoding();
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            // Filename e.g IFACEFIXED
            Match matchOne = Regex.Match(fileName, @"[a-zA-Z]+-[0-9]{1,10}(\w+)");
            string fileType = matchOne.Groups[1].Value.Replace("PC", "").Replace("APPS", "");
            fileType = Regex.Replace(fileType, @"\d+$", "");
            filesEncoding.FileType = fileType;
            // Filename with -NP & -P
            if (fileName.Contains("-NP") || fileName.Contains("-P"))
            {
                string charPattern = @"-(NP|P)";
                Match matchTwo = Regex.Match(fileName, charPattern);
                string characterType = matchTwo.Groups[1].Value;
                if (characterType == "NP")
                    filesEncoding.CharacterType = characterType.Replace(characterType.ToString(), "Non-Polish");
                else if (characterType == "P")
                    filesEncoding.CharacterType = characterType.Replace(characterType.ToString(), "Polish");

                if (fileName.Contains("PC"))
                    filesEncoding.From = "PC";
                else if (fileName.Contains("APPS"))
                    filesEncoding.From = "Server";
                else
                    filesEncoding.From = "Email";
            }
            // Not based with Filename
            else
            {
                bool containsPolishCharacters = ContainsPolishCharacters(ReadingOfFiles(filePath));
                if (containsPolishCharacters)
                    filesEncoding.CharacterType = "Polish";
                else if (!containsPolishCharacters)
                    filesEncoding.CharacterType = "Non-Polish";
            }
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                /// <summary>
                /// UTF8    : EF BB BF
                /// UTF16 BE: FE FF
                /// UTF16 LE: FF FE
                /// UTF32 BE: 00 00 FE FF
                /// UTF32 LE: FF FE 00 00
                /// UDE -----
                /// UTF8 BOM = UTF8
                /// UTF16 BOM = UTF16
                /// UTF8 = ASCII
                /// UTF16 LE = WINDOWS 1252
                /// </summary>

                // Read the first few bytes from the file
                byte[] buffer = new byte[4];
                fs.Read(buffer, 0, 4);

                bool isUtf8 = false, isUtf16Le = false; //isUtf8 (true) = with bom | isUtf8 (false) = no bom | isUtf16Le (true) = no bom | isUtf16Le (false) = with bom

                var utf8NoBom = new UTF8Encoding(false);
                using (var reader = new StreamReader(filePath, utf8NoBom))
                {
                    reader.Read(); //identify bom || !bom

                    if (reader.CurrentEncoding == utf8NoBom) // !BOM
                    {
                        Ude.CharsetDetector cdet = new Ude.CharsetDetector();
                        cdet.Feed(fs);
                        cdet.DataEnd();

                        if (buffer[0] == 0x30 && buffer[1] == 0 && buffer[2] == 0x30 && buffer[3] == 0)
                            filesEncoding.EncodingType = "UTF-16 LE";
                        else if (cdet.Charset == "ASCII" || cdet.Charset == "UTF-8")
                            filesEncoding.EncodingType = "UTF-8";
                        else if (cdet.Charset == "windows-1252" || cdet.Charset == "UTF-16 LE")
                            filesEncoding.EncodingType = "UTF-16 LE";
                        else
                            filesEncoding.EncodingType = "Undetected Encoding";
                    }
                    else // BOM
                    {
                        if (buffer.Length >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
                            filesEncoding.EncodingType = "UTF-8 BOM";
                        else if (buffer.Length >= 2 && buffer[0] == 0xFF && buffer[1] == 0xFE)
                            filesEncoding.EncodingType = "UTF-16 LE BOM";
                        else
                            filesEncoding.EncodingType = "Undetected Encoding";
                    }
                }
                filesEncoding.FileName = fileName;
                filesEncodings.Add(filesEncoding);
            }
            return null;
        }
        public void MatchList(StreamWriter sw)
        {
            int countPassed = 0, countFailed = 0;
            for (int j = 0; j < filesEncodings.Count; j++)
            {
                bool isMatch = false;
                var data = filesEncodings[j];
                for (int i = 0; i < baseFileEncoding.Count; i++)
                {
                    isMatch = false;
                    var baseData = baseFileEncoding[i];

                    foreach (var source in baseData.Sources)
                    {
                        if (data.FileType == baseData.FileType && data.CharacterType == baseData.CharacterType && data.EncodingType == source.EncodingType) // MATCH / PASSED
                        {
                            if(data.From == source.From)
                            {
                                sw.WriteLine($"{data.FileName} [{data.CharacterType} | {data.EncodingType} | {data.FileType} | {data.From}] - PASSED");
                                isMatch = true;
                                countPassed++;
                                break; // Exit the inner loop if a match is found
                            }
                            else if (string.IsNullOrEmpty(data.From))
                            {
                                sw.WriteLine($"{data.FileName} [{data.CharacterType} | {data.EncodingType} | {data.FileType } | NULL ] - PASSED");
                                isMatch = true;
                                countPassed++;
                                break; // Exit the inner loop if a match is found
                            }
                        }
                    }
                    if (isMatch == true)
                        break;
                }
                if (isMatch == false)
                {
                    sw.WriteLine($"{data.FileName} [{data.CharacterType} | " +
                                $"{data.EncodingType} | {data.FileType} | {(data.From != null ? data.From : "Undetected Source")} ] - FAILED");
                    countFailed++;
                }
            }
            sw.WriteLine($"Count Passed: {countPassed}\nCount Failed: {countFailed}");
            sw.Close();
        }
        static bool ContainsPolishCharacters(string text)
        {
            foreach (char c in text)
                if (IsPolishDiacritic(c))
                    return true;

            return false;
        }
        static bool IsPolishDiacritic(char c)
        {
            // Polish diacritics range in Unicode
            // You can adjust this range if needed
            Regex polishDiacriticsRegex = new Regex(@"[ąćęłńóśźżĄĆĘŁŃÓŚŹŻßÄäÖöÜüÀàÂâÉéÈèÊêËëÎîÏïÔôŒœÙùÛûÙúÑñÜü]");
            return polishDiacriticsRegex.IsMatch(c.ToString(CultureInfo.InvariantCulture));
        }
        private static string ReadingOfFiles(string file)
        {
            string text = File.ReadAllText(file);
            return text;
        }
        public void ListOfEncoding()
        {
            FileInfo fiOne = new FileInfo(ConfigurationSettings.AppSettings["EncodingBase"]);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelOne = new ExcelPackage(fiOne);
            var wsOne = excelOne.Workbook.Worksheets[0];
            int row = wsOne.Dimension.End.Row;
            int col = wsOne.Dimension.End.Column;
            for (int i = 1; i <= row; i++)
            {
                for (int j = 1; j <= col; j++) // A/1-FileType, C/3-Server, D/4-Email, E/5-PC
                {
                    if(i >= 3 && i <= 27) // Non - Polish
                    {
                        if (i % 2 != 0 && j != 2)
                        {
                            var dataEncoding = new BaseFileEncoding
                            {
                                FileType = wsOne.Cells[i, 1].Text,
                                Sources = new List<BaseSource>
                                {
                                    new BaseSource { From = "Server", EncodingType = wsOne.Cells[i, 3].Text },
                                    new BaseSource { From = "Email", EncodingType = wsOne.Cells[i, 4].Text },
                                    new BaseSource { From = "PC", EncodingType = wsOne.Cells[i, 5].Text }
                                },
                                CharacterType = "Non-Polish"
                            };
                            baseFileEncoding.Add(dataEncoding);
                            break;
                        }
                    }
                    else if (i >= 33 && i <= 57) // Polish
                    {
                        if(i % 2 != 0 && j != 2) 
                        {
                            var dataEncoding = new BaseFileEncoding
                            {
                                FileType = wsOne.Cells[i, 1].Text,
                                Sources = new List<BaseSource>
                                {
                                    new BaseSource { From = "Server", EncodingType = wsOne.Cells[i, 3].Text },
                                    new BaseSource { From = "Email", EncodingType = wsOne.Cells[i, 4].Text },
                                    new BaseSource { From = "PC", EncodingType = wsOne.Cells[i, 5].Text }
                                },
                                CharacterType = "Polish"
                            };
                            baseFileEncoding.Add(dataEncoding); 
                            break;
                        }
                    } 
                }
            }
        }
        List<FilesEncoding> filesEncodings = new List<FilesEncoding>();
        List<BaseFileEncoding> baseFileEncoding = new List<BaseFileEncoding>();
        public class FilesEncoding
        {
            public string FileName { get; set; }
            public string FileType { get; set; }
            public string EncodingType { get; set; }
            public string From { get; set; }
            public string CharacterType { get; set; }
        }
        public class BaseFileEncoding
        {
            public string FileType { get; set; }
            public List<BaseSource> Sources { get; set; }
            public string CharacterType { get; set; }
        }
        public class BaseSource
        {
            public string From { get; set; }
            public string EncodingType { get; set; }
        }
    }
}
