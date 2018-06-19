using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       
using System.Threading;

namespace Utilities
{
    public class ReadExcelFile
    {
        /*
        public ReadExcelFile()
        {

        }
        public ReadExcelFile (string ExcelFileAndPath)
        {

        }
        */
        //public static void getExcelFile()
        //List<string> ExcelList = new List<string>();
        List<string> ExcelList = new List<string>();
        List<string> ExcelAdditiveList = new List<string>();
        List<string> ExcelAdditiveListNoNull = new List<string>();

        public List<string> getExcelFile(string ExcelFileAndPath, int WorkSheetNumber) //plan on overloading this method again to pass in desired sheets or name

        {
            //Create COM Objects. Create a COM object for everything that is referenced

            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelFileAndPath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[WorkSheetNumber];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            string ExcelTemp = string.Empty;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                    {
                        //Console.Write("\r\n");
                        //Console.Write("Beep\n");
                        
                        //ExcelList.Add(xlRange.Cells[i, j].Value2.ToString() + "\t");
                    }

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                    
                        //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "|");
                        //Thread.Sleep(1000);

                        ExcelList.Add(xlRange.Cells[i, j].Value2.ToString() + "|");

                    }

                } //column count complete


            }


            int kCount = 0;
            int kcolCount = 0;
            int counter = 0;

            for (int i = 0; i < rowCount; i++)
            {
                {
                    string _temp = "";
                    string _temp_ = "";
                    kCount = counter;
                    kcolCount = colCount + kCount;
                    //int firstIndex = rowCount / colCount;

                    for (int k = counter; k < kcolCount; k++)
                    {
                        Console.Write(ExcelList[k]);
                        Thread.Sleep(250);

                        _temp_ = _temp_ + ExcelList[k].ToString();
                        counter++;
                    }
                    ExcelAdditiveList.Add(_temp_);
                    //kCount = counter;
                    //kcolCount = colCount + kCount;
                    Console.WriteLine();
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            
            foreach (string temp in ExcelAdditiveList)
            {
                string ExcelNoNull = Substitute.ReplaceReplace(temp, "NULL", "");
                ExcelAdditiveListNoNull.Add(ExcelNoNull);
            }

            return ExcelAdditiveListNoNull;
        }
                 
    }
}
