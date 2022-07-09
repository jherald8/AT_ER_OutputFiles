using System;

namespace QATest
{
    class Program
    {
        static void Main(string[] args)
        {
            TextComparison textComparison = new TextComparison();
            textComparison.CompareTxtFiles();
            ExcelComparison excelComparison = new ExcelComparison();
            excelComparison.CompareExcelFiles();
            CSVComparison csvComparison = new CSVComparison();
            csvComparison.CompareCSVFiles();
        }
    }
}
