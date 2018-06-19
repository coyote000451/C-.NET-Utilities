using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;

namespace ReadingAccessDB
{
    class DumpDataSetToConsole
    {
        public static void DumpDataSet(DataSet ds)
        {
            Console.Out.WriteLine("DataSet: {0}", ds.DataSetName);

            // For every tables in the DataSet ...
            foreach (DataTable dt in ds.Tables)
            {
                Console.Out.WriteLine("\tTableName: {0}", dt.TableName);

                // ... Write the table schema
                foreach (DataColumn col in dt.Columns)
                {
                    Console.Out.Write("\t\t" + col.ColumnName + " ");
                }
                Console.Out.WriteLine("\t\t");

                // ... Write the table contents
                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        Console.Out.Write("\t\t" + row[i]);
                    }
                    Console.Out.WriteLine("");
                }
            }
        }
    }
}
