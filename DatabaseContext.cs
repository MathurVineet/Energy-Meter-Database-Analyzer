using System;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;

namespace DbExport
{
    public static class DatabaseContext
    {
        public static SQLiteConnection InitializeConnection()
        {
            try
            {
                //Initialize Connection
                string connString = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UserMasterDb.db");
                SQLiteConnection conn = new SQLiteConnection($"Data Source={connString};Version=3;New=True;Compress=True;");
                return conn;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error in conn", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }

        //public bool ExecuteSelectCommand(string command)
        //{
        //    return false;
        //}
        public static object ExecuteScallerValue(string command)
        {
            try
            {
                //Initialize connection and open
                SQLiteConnection conn = InitializeConnection();
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand(command, conn);
                object returnValue = cmd.ExecuteScalar();
                conn.Close();
                conn.Dispose();
                return returnValue;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public static int ExecuteNonQuery(string command)
        {
            try
            {
                //Initialize connection and open
                SQLiteConnection conn = InitializeConnection();
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand(command, conn);
                int returnValue = Convert.ToInt32(cmd.ExecuteNonQuery());
                conn.Close();
                conn.Dispose();
                return returnValue;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }

        }
        public static int ExecuteScaller(string command)
        {
            try
            {
                //Initialize connection and open
                SQLiteConnection conn = InitializeConnection();
                conn.Open();

                SQLiteCommand cmd = new SQLiteCommand(command, conn);
                int returnValue = Convert.ToInt32(cmd.ExecuteScalar());
                conn.Close();
                conn.Dispose();
                return returnValue;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
             
        }

       
    }
}
