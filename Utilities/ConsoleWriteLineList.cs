using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadingExcel
{
    class ConsoleWriteLineList
    {
        public static void DumpExcelSet(List<string> dataSet)
        {
            foreach (string e in dataSet)
                Console.Write("{0}", e);

        }
    }
}
