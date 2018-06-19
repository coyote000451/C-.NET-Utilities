using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Runtime.InteropServices;
using ReadingAccessDB;

namespace Utilities
{
    class ReadExcelFileReturnList
    { 

    List<string> ExcelList = new List<string>();
    List<string> ExcelAdditiveList = new List<string>();
    List<string> ExcelAdditiveListNoNull = new List<string>();
        //string[,] ExcelArray;

    public List<string> getExcelFile(string ExcelFileAndPath, int WorkSheetNumber) //plan on overloading this method again to pass in desired sheets or name

    {
        //Create COM Objects. Create a COM object for everything that is referenced

        Excel.Application xlApp = new Excel.Application();

        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelFileAndPath);
        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[WorkSheetNumber];
        Excel.Range xlRange = xlWorksheet.UsedRange;

        int rowCount = xlRange.Rows.Count;
        int colCount = xlRange.Columns.Count;
        string[,] ExcelTemp = new string[rowCount, colCount];


        //iterate over the rows and columns and print to the console as it appears in the file
        //excel is not zero based!!
        for (int i = 1; i <= rowCount; i++)
        {
            for (int j = 1; j < colCount; j++)
            {

                //write the value to the console
                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                {
                    ExcelTemp[i,j] = xlRange.Cells[i, j].ToString();
                }

            } //column count complete


        }

        //cleanup
        GC.Collect();
        GC.WaitForPendingFinalizers();

        Marshal.ReleaseComObject(xlRange);
        Marshal.ReleaseComObject(xlWorksheet);

        xlWorkbook.Close();
        Marshal.ReleaseComObject(xlWorkbook);

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
