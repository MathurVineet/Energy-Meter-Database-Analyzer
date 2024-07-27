using System;
using System.Data.SQLite;

namespace DbExport
{
    public class CommonUtility
    {
        public string GetReportFilePath()
        {
            try
            {
                //SQLiteConnection conn = DatabaseContext.InitializeConnection();
                
                string Sql = "select mdbFilePath from FilePaths";
                object result = DatabaseContext.ExecuteScallerValue(Sql);
               
                if (result != null)
                    return result.ToString();
                else
                    return null;  // Or handle this case as needed                
            }
            catch (Exception)
            {
                throw;
            }

        }
    }
}
