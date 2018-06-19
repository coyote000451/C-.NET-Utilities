//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Access;
using System.Data;
using System.Data.OleDb;

namespace Utilities
{
    public static class ExportAccessTabletoExcel
    {
        public static void ExportQuery(string databaseLocation, string tableNameToExport, string excelLocationToExportTo)
        {
            var application = new Application();
            application.OpenCurrentDatabase(databaseLocation);

            ///acSpreadsheetTypeExcel12Xml is coincident with the version of Excel installed on machine
            application.DoCmd.TransferSpreadsheet(AcDataTransferType.acExport, AcSpreadSheetType.acSpreadsheetTypeExcel12Xml,
                                                  tableNameToExport, excelLocationToExportTo, true);
            application.CloseCurrentDatabase();
            application.Quit();
            Marshal.ReleaseComObject(application);

        
        }
    }

    //ExportQuery(@"C:\blah\blah.accdb", "AccessTableName", @"C:\<PATH_And_Name_To_Excel_File.xlsx>");
}
