using System;
using System.IO;
using System.Linq;
using System.Text;

namespace QATestEncoding
{
    internal class Program
    {
        static void Main(string[] args)
        {
            encodingTypes encoding =  new encodingTypes();
            encoding.DecryptingExcel(@"C:\Users\jmartin\Downloads\Jerald Files\DailyTask Test\QATest\");
        }
    }
}
