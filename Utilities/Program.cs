using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReadingAccessDB;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;

namespace ReadingExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            ExportAccessTabletoExcel.ExportQuery(@"C:\Projects\CHIA\Cerner\Cerner_Multum_en-US_249_180601.accdb", "drc_gestational_age_non_continuous_infusion", @"C:\Projects\CHIA\drc_gestational_age_non_continuous_infusion.xlsx");

            ReadExcelFileReturnList GetGetExcel = new ReadExcelFileReturnList();
            GetGetExcel.getExcelFile(@"C:\Projects\CHIA\drc_gestational_age_non_continuous_infusion.xlsx", 1);

            //List<string> ExcelList = new List<string>();

            //ReadExcelFile GetExcel = new ReadExcelFile(@"C:\Projects\CHIA\Copyof20180513_drc_age2_updates_US_CA_v249.xlsx");

            ReadExcelFile GetExcel = new ReadExcelFile();
            //ReadExcelFile GetExcel = new ReadExcelFile();
            //GetExcel.ExcelList = ReadExcelFile.getExcelFile();


            List<string> ExcelList = GetExcel.getExcelFile(@"C:\Projects\CHIA\Copyof20180513_drc_age2_updates_US_CA_v251.xlsx", 3); //Pass in Excel path and string
            //List<string> ExcelList2 = GetExcel2.getExcelFile(@"C:\Projects\CHIA\Copyof20180513_drc_age2_updates_US_CA_v251.xlsx", 3);
            List <string> ExcelList2 = GetExcel.getExcelFile(@"C:\Projects\CHIA\drc_gestational_age_non_continuous_infusion.xlsx", 1);

            CompareList CompareTwoExcelLists = new CompareList();
            //List<string> ReturnLists = CompareTwoLists.Contains(ExcelList, DataList);
            IEnumerable<string> ReturnDiffLists = CompareTwoExcelLists.Contains(ExcelList, ExcelList2);
            PrintToFile.Print(ExcelList);
            PrintToFile.Print(ExcelList2);
            PrintToFile.Print(ReturnDiffLists);


            //ConsoleWriteLineList.DumpExcelSet(ExcelList);

            // Load an Access ACCDB file

            DataSet ds1 = AccessDbLoader.LoadFromFile(@"C:\Projects\CHIA\Cerner\Cerner_Multum_en-US_249_180601.accdb");
            //DataSet ds1 = AccessDbLoader.LoadFromFile(@"C:\Projects\CHIA\Cerner\Cerner_Multum_en-CA_249_180601.accdb");
            //DataSet ds1 = AccessDbLoader.LoadFromFile(@"C:\Projects\CHIA\Lexi\Cerner_LexiComp_en-US_249_180601.accdb");
            List<string> DataList = DataSetToList.DumpDataSet(ds1);

            //DumpDataSetToConsole.DumpDataSet(ds1);

            //var firstNotSecond = ExcelList.Except(DataList).ToList();
            //var secondNotFirst = DataList.Except(ExcelList).ToList();

            CompareList CompareTwoLists = new CompareList();
            //List<string> ReturnLists = CompareTwoLists.Contains(ExcelList, DataList);
            IEnumerable<string> ReturnLists = CompareTwoLists.Contains(ExcelList, DataList);
            PrintToFile.Print(ExcelList);
            PrintToFile.Print(DataList);
            PrintToFile.Print(ReturnLists);
            //ConsoleWriteLineList.DumpExcelSet(ReturnLists);
            Console.ReadLine();
        }
    }
}
