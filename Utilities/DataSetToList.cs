using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.IO;
//using System.Data.OleDb;
using System.Data;
using System.Threading;
using ReadingAccessDB;

namespace Utilities
{
    class DataSetToList
    {
        public static List<string> DumpDataSet(DataSet ds)
        {
            //////List<string> DumpDataList = new List<string>();
            List<DataRow> DumpDataList = new List<DataRow>();
            List<string> DumpDataAdditiveList = new List<string>();
            List<string> CombinedRowList = new List<string>();
            int RowCount = 0;
            int ColumnCount = 0;

            // For every tables in the DataSet ...
            foreach (DataTable dt in ds.Tables)
            {
                Console.WriteLine("Table:  {0}", dt);
                // ... Write the table contents
                List<string> tempList = new List<string>();
                int count = 0;

                foreach (DataRow row in dt.Rows)
                {                    
                    DumpDataList.Add(row);                    
                }

                ColumnCount = dt.Columns.Count;
                RowCount = dt.Rows.Count;

                foreach (DataRow temp in DumpDataList)
                {
                    string _temp = "";
                    string _temp_ = "";
                    for (int k = 0; k < ColumnCount; k++)
                    {
                        Console.Write(temp.ItemArray[k] + "|");

                        _temp_ = _temp_ + "|" + temp.ItemArray[k].ToString();
                        _temp = _temp_;
                    }
                    CombinedRowList.Add(_temp_);
                    Console.WriteLine();
                }

                Console.WriteLine("Row Count Total: {0}", count);
                Thread.Sleep(10000);
            }
            //return DumpDataAdditiveList;
            foreach (string found in CombinedRowList)
            {
                string getRidofFirstPipe = Substitute.ReplaceFirstOccurrence(found, "|", "");
                DumpDataAdditiveList.Add(getRidofFirstPipe);
            }
            
            return DumpDataAdditiveList;
        }

    }


}
