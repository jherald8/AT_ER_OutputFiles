using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QATest
{
    class TextComparison
    {
        public void CompareTxtFiles()
        {
            string[] pathA = Directory.GetFiles(@"D:\Work\TestQA\pathA\", "*txt");
            string[] pathB = Directory.GetFiles(@"D:\Work\TestQA\pathB\", "*txt");

            foreach (var oneFile in pathA)
            {
                foreach (var twoFile in pathB)
                {
                    if (Path.GetFileName(oneFile) == Path.GetFileName(twoFile))
                    {
                        string[] lines = File.ReadAllLines(oneFile);
                        string[] lines2 = File.ReadAllLines(twoFile);
                        string passedDest = @"D:\Work\TestQA\Passed\";
                        string failedDest = @"D:\Work\TestQA\Failed\";
                        bool passed = false;
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
                            File.Move(oneFile, passedDest + Path.GetFileName(oneFile));
                        }
                        else
                        {
                            File.Move(oneFile, failedDest + Path.GetFileName(oneFile));
                        }
                        break;
                    }
                    else
                    {

                    }
                }
            }
        }
    }
}

