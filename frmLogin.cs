using System;
using System.Windows.Forms;

namespace DbExport
{
    public partial class frmLogin : Form
    {
        //private SQLiteConnection sqlite_conn;
        public frmLogin()
        {
            InitializeComponent();    
        }             
        
        private void btnLogin_Click(object sender, EventArgs e)
        {
            //Condition for check empty field
            if (string.IsNullOrEmpty(txtUserName.Text.Trim()))
            {
                MessageBox.Show("Required Username.", "Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtUserName.Focus();
                return;
            }

            if (string.IsNullOrEmpty(txtPassword.Text.Trim()))
            {
                MessageBox.Show("Required Password.", "Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtPassword.Focus();
                return;
            }

            //Initialize Select statement
            string query = "select count(*) from UserMaster where UserName = '" +
                           txtUserName.Text.Trim() + "' and PassWord = '" +
                           txtPassword.Text.Trim() + "'";

            //Call method for execute query
            int returnValue = DatabaseContext.ExecuteScaller(query);

            if (returnValue <= 0)
            {
                MessageBox.Show("Invalid Username or Password.");
                txtUserName.Focus();
                return;
            }

            //Open main form
            frmReports frmReports = new frmReports();
            frmReports.ShowDialog();
        }
          
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
