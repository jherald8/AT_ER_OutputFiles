using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace QATest
{
    class ExcelComparison
    {
        public void CompareExcelFiles()
        {
            string[] pathA = Directory.GetFiles(@"C:\Users\jmartin\Downloads\Jerald Files\DailyTask Test\QATest\Excel\", "*xlsx");
            string[] pathB = Directory.GetFiles(@"C:\Users\jmartin\Downloads\Jerald Files\DailyTask Test\QATest\Excel\current\", "*xlsx");
            string failedDest = @"D:\Work\TestQA\ExcelFile\Failed\";
            foreach (var fileOne in pathA)
            {
                foreach (var fileTwo in pathB)
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
                        for (int i = 1; i <= row; i++)
                        {
                            for (int j = 1; j <= col; j++)
                            {
                                if (wsOne.Cells[i, j].Value.ToString() == wsTwo.Cells[i, j].Value.ToString())
                                {

                                }
                                else
                                {
                                    using (ExcelRange rng = wsOne.Cells[i, j])
                                    {
                                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        rng.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                                    }
                                }
                            }
                        }
                        excelOne.Save();
                        break;
                    }
                    else
                    {
                        File.Move(fileOne, failedDest + Path.GetFileName(fileOne));
                    }
                }
            }
        }
    }
}
