using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QATest
{
    class CSVComparison
    {
        public void CompareCSVFiles()
        {
            using (StreamReader f1 = new StreamReader(@"D:\Work\TestQA\CSVFile\pathA\OPSH-924COMMA20220617221056.csv"))
            using (StreamReader f2 = new StreamReader(@"D:\Work\TestQA\CSVFile\pathB\OPSH-924COMMA20220617221056.csv"))
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
                        differences.Add(string.Format("Line {0} differs. File 1: {1}, File 2: {2}", lineNumber, line1, line2));
                    }
                }

                if (!f2.EndOfStream)
                {
                    differences.Add("Differing number of lines - f1 has less.");
                }
            }

            //    string[] firstCSV = System.IO.File.ReadAllLines(@"D:\Work\TestQA\CSVFile\pathA\OPSH-924COMMA20220617221056.csv");
            //    string[] secondCSV = System.IO.File.ReadAllLines(@"D:\Work\TestQA\CSVFile\pathB\OPSH-924COMMA20220617221056.csv");

            //    // Create the query. Note that method syntax must be used here.
            //    IEnumerable<string> szDifference =
            //    firstCSV.Except(secondCSV);

            //    foreach (string szTest in szDifference)
            //        Console.WriteLine(szTest + " exist in firstCSV but not in secondCSV");
        }
    }
}
