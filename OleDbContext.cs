using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace DbExport
{
    public static class OleDbContext
    {
        private static OleDbConnection _connection;
        private static string _connectionString;
        public static OleDbConnection InitializeConnection(string dbfilePath)
        {
            try {
                //Initialize Connection
                _connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbfilePath + ";Jet OLEDB:Database Password=redware@12345;";
                _connection = new OleDbConnection(_connectionString);
                _connection.Open();
                return _connection;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error in conn", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            
        }
       
        public static OleDbDataReader ExecuteReader(string query)
        {
            OleDbCommand cmd = new OleDbCommand(query, _connection);
            return cmd.ExecuteReader();
        }

       
        public static DataSet ExecuteQuery(string query,string dbfilePath)
        {
            try
            {
                OleDbConnection conn = InitializeConnection(dbfilePath);
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText = query;
                cmd.Connection = conn;
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                conn.Close();
                conn.Dispose();
                return ds;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }
        
    }


}
