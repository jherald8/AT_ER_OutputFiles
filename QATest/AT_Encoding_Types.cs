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
            string source = ConfigurationSettings.AppSettings["SourcePath"];
            FileTool fileTool = new FileTool();
            string[] newFiles = Directory.GetFiles(source);
            string logFile = $@"c:\temp\EncodingLog{DateTime.Now.ToString("yyyyMdHHmmss")}.txt";
            StreamWriter sw;
            sw = File.CreateText(logFile);

            foreach (string file in newFiles)
                if (Path.GetExtension(file) == ".zip" || Path.GetExtension(file) == ".ZIP")
                {
                    fileTool.Decompress(source);
                    break;
                }
            foreach (var file in Directory.GetFiles(source))
                DetectFileEncoding(file);
            MatchList(sw);
        }
        public Encoding DetectFileEncoding(string filePath)
        {
            FilesEncoding filesEncoding = new FilesEncoding();
            string fileName = Path.GetFileNameWithoutExtension(filePath);

            // Filename e.g IFACEFIXED
            string ftPattern = @"OPSH-9322(\w+)";
            Match matchOne = Regex.Match(fileName, ftPattern);
            string fileType = matchOne.Groups[1].Value;
            fileType = Regex.Replace(fileType, @"\d+$", "");
            filesEncoding.FileType = fileType;

            if (fileName.Contains("-NP") || fileName.Contains("-P"))
            {
                // Filename with -NP & -P
                string charPattern = @"-(NP|P)";
                Match matchTwo = Regex.Match(fileName, charPattern);
                string characterType = matchTwo.Groups[1].Value;
                if (characterType == "NP")
                    filesEncoding.CharacterType = characterType.Replace(characterType.ToString(), "Non-Polish");
                else if (characterType == "P")
                    filesEncoding.CharacterType = characterType.Replace(characterType.ToString(), "Polish");
            }
            else
            {
                ReadingOfFiles(filePath);
                bool containsPolishCharacters = ContainsPolishCharacters(ReadingOfFiles(filePath));
                if (containsPolishCharacters)
                    filesEncoding.CharacterType = "Polish";
                else if (!containsPolishCharacters)
                    filesEncoding.CharacterType = "Non-Polish";
            }
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Read the first few bytes from the file
                byte[] buffer = new byte[4];
                fs.Read(buffer, 0, 4);

                var utf8NoBom = new UTF8Encoding(false);
                using (var reader = new StreamReader(filePath, utf8NoBom))
                {
                    reader.Read();
                    if (Equals(reader.CurrentEncoding, utf8NoBom))
                        filesEncoding.EncodingType = "UTF-8";
                    else
                    {
                        //WITH BOM
                        if (buffer.Length >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
                            filesEncoding.EncodingType = "UTF-8 BOM";
                        else if (buffer.Length >= 2 && buffer[0] == 0xFF && buffer[1] == 0xFE)
                            filesEncoding.EncodingType = "UTF-16 LE BOM";
                    }
                }
                filesEncoding.FileName = fileName;
                filesEncodings.Add(filesEncoding);
            }
            return null;
        }
        public void MatchList(StreamWriter sw)
        {
            foreach (var file in filesEncodings)
                foreach (var fileList in listOfEncoding)
                {
                    if (file.CharacterType == fileList.CharacterType && file.EncodingType == fileList.EncodingType && file.EncodingType == fileList.EncodingType)
                    {
                        sw.WriteLine($"{file.FileName} [{file.CharacterType} | " +
                            $"{file.EncodingType} | {file.FileType}] - PASSED");
                        break;
                    }
                    else if (file.FileType == fileList.FileType && file.CharacterType == fileList.CharacterType && file.EncodingType != fileList.EncodingType)
                        sw.WriteLine($"{file.FileName} [{file.CharacterType} | " +
                            $"{file.EncodingType} | {file.FileType}] - FAILED");
                }
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
        public class FilesEncoding
        {
            public string FileName { get; set; }
            public string FileType { get; set; }
            public string EncodingType { get; set; }
            public string CharacterType { get; set; }
        }
        List<FilesEncoding> filesEncodings = new List<FilesEncoding>();
        public class FileEncoding
        {
            public string FileType { get; set; }
            public string EncodingType { get; set; }
            public string CharacterType { get; set; }
        }
        List<FileEncoding> listOfEncoding = new List<FileEncoding>
            {
                // POLISH
                new FileEncoding{ FileType = "COMMA", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "COMMAH", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "XLS", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "EXCRAW", EncodingType = "UTF-16 LE BOM", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "EXCURR", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "EXFMT", EncodingType = "UTF-16 LE BOM", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "IFACECD", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "IFACEDEL", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "IFACEFIXED", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "IFACETAB", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "SCOLON", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "SCOLONH", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                new FileEncoding{ FileType = "TXT", EncodingType = "UTF-8", CharacterType = "Non-Polish"},
                // NON-POLISH
                new FileEncoding{ FileType = "COMMA", EncodingType = "UTF-8 BOM", CharacterType = "Polish"},
                new FileEncoding{ FileType = "COMMAH", EncodingType = "UTF-8 BOM", CharacterType = "Polish"},
                new FileEncoding{ FileType = "XLS", EncodingType = "UTF-16 LE BOM", CharacterType = "Polish"},
                new FileEncoding{ FileType = "EXCRAW", EncodingType = "UTF-16 LE BOM", CharacterType = "Polish"},
                new FileEncoding{ FileType = "EXCURR", EncodingType = "UTF-16 LE BOM", CharacterType = "Polish"},
                new FileEncoding{ FileType = "EXFMT", EncodingType = "UTF-16 LE BOM", CharacterType = "Polish"},
                new FileEncoding{ FileType = "IFACECD", EncodingType = "UTF-8", CharacterType = "Polish"},
                new FileEncoding{ FileType = "IFACEDEL", EncodingType = "UTF-8", CharacterType = "Polish"},
                new FileEncoding{ FileType = "IFACEFIXED", EncodingType = "UTF-8", CharacterType = "Polish"},
                new FileEncoding{ FileType = "IFACETAB", EncodingType = "UTF-8", CharacterType = "Polish"},
                new FileEncoding{ FileType = "SCOLON", EncodingType = "UTF-8 BOM", CharacterType = "Polish"},
                new FileEncoding{ FileType = "SCOLONH", EncodingType = "UTF-8 BOM", CharacterType = "Polish"},
                new FileEncoding{ FileType = "TXT", EncodingType = "UTF-8 BOM", CharacterType = "Polish"},
            };
    }
}
