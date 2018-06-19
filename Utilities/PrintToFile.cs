using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Utilities
{
    public class PrintToFile
    {
        public static void Print(IEnumerable<string> fileName)
        {

            string GeneratedFileName = Path.GetRandomFileName();
            using (StreamWriter outputFile = new StreamWriter(@"c:\Projects\CHIA\OutPut_" + GeneratedFileName + ".txt"))
                //while ((line = file.ReadLine()) != null)

                foreach (string e in fileName)
                {
                    // Console.Write("{0}", e);
                    outputFile.WriteLine(e);
                }

        }
    }
}
