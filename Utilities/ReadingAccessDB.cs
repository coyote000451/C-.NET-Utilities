
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace Utilities

{
    /// <summary>
    /// Useful utilities for Microsoft Access Database files.
    /// </summary>
    public static class AccessDbLoader
    {

        public static DataSet LoadFromFile(string fileName)
        {
            DataSet result = new DataSet();

            // For convenience, the DataSet is identified by the name of the loaded file (without extension).
            result.DataSetName = Path.GetFileNameWithoutExtension(fileName).Replace(" ", "_");

            // Compute the ConnectionString (using the OLEDB v12.0 driver compatible with ACCDB and MDB files)
            fileName = Path.GetFullPath(fileName);
            string connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};User Id=Admin;Password=", fileName);

            // Opening the Access connection
            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                conn.Open();

                // Getting all user tables present in the Access file (Msys* tables are system thus useless for us)
                DataTable dt = conn.GetSchema("Tables");
                List<string> tablesName = dt.AsEnumerable().Select(dr => dr.Field<string>("TABLE_NAME")).Where(dr => !dr.StartsWith("MSys")).ToList();

                // Getting the data for every user tables
                //foreach (string tableName in tablesName)
                //string tableName = "drc_extract_advanced";
                //string tableName = "groups_obsolete";
                //string tableName = "drc_extract_patient_specific";
                string tableName = "drc_gestational_age_non_continuous_infusion";
                //{
                using (OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}]", tableName), conn))
                    {
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                        {
                            // Saving all tables in our result DataSet.
                            DataTable buf = new DataTable("[" + tableName + "]");
                            adapter.Fill(buf);
                            result.Tables.Add(buf);
                        } // adapter
                    } // cmd
                //} // tableName

            } // conn

            // Return the filled DataSet
            return result;
        }
    }
}
