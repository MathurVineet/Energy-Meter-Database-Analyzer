using ClosedXML.Excel;
using DocumentFormat.OpenXml.Vml;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;


namespace DbExport
{
    public partial class frmReports : Form
    {
        string export_key_Sitzungsnr = "";
        string dbfilePath = "";
        //string exportFileName = "";
        public frmReports()
        {
            InitializeComponent();
            btnExportToExcel.Enabled = false;
        }    
        private void PopulateDataGrid(DataSet ds)
        {
           
            dataGridReport.DataSource = null;
            dataGridReport.Columns.Clear();
            dataGridReport.MultiSelect = false;
            dataGridReport.DataSource = ds.Tables[0];
            // ---columns headings--
            dataGridReport.Columns.Add("SerialNumber", "S.No.");
            dataGridReport.Columns["PAName"].HeaderText = "PA Name";
            dataGridReport.Columns["Datum"].HeaderText = "Date";
            dataGridReport.Columns["Beginn"].HeaderText = "Start";
            dataGridReport.Columns["Ende"].HeaderText = "End";
            dataGridReport.Columns["Erstelldatum"].HeaderText = "Creation Date";
            dataGridReport.Columns["Sitzungsnr"].HeaderText = "Sitzungsnr";

            //---columns data -----
            // Generate the serial numbers
            for (int i = 0; i < dataGridReport.Rows.Count; i++)
                dataGridReport.Rows[i].Cells["SerialNumber"].Value = (i + 1).ToString();

            // Move the serial number column to the first position
            dataGridReport.Columns["SerialNumber"].DisplayIndex = 0;

            // Set the width of the columns explicitly
            dataGridReport.Columns["SerialNumber"].Width = 50;
            dataGridReport.Columns["PAName"].Width = 350;
            dataGridReport.Columns["Datum"].Width = 150;
            dataGridReport.Columns["Beginn"].Width = 150;
            dataGridReport.Columns["Ende"].Width = 150;
            dataGridReport.Columns["Erstelldatum"].Width = 150;

            // Hide the Sitzungsnr column
            dataGridReport.Columns["Sitzungsnr"].Visible = false;

            dataGridReport.ClearSelection();
        }

        private void frmReports_Load(object sender, EventArgs e)
        {
            btnExportToExcel.Enabled = false;

            CommonUtility commonUtility = new CommonUtility();
            dbfilePath = commonUtility.GetReportFilePath();

            if (string.IsNullOrEmpty(dbfilePath) || !File.Exists(dbfilePath))
            {
                MessageBox.Show("File Path not available. please select a DB file");
                btnChangeDatabase.Focus();
                return;
            }
            string sql = "SELECT PAName,Datum, Beginn,Ende, Erstelldatum,Sitzungsnr  FROM Report;";
            DataSet ds = OleDbContext.ExecuteQuery(sql, dbfilePath);
            PopulateDataGrid(ds);
            
        }

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection conn = OleDbContext.InitializeConnection(dbfilePath);
                
                //string query = "select * from RepSeriennr where Kennung Like \"Sr number\" and Sitzungsnr = " + export_key_Sitzungsnr + "; ";
                string query = "SELECT * FROM RepSeriennr WHERE(((RepSeriennr.Sitzungsnr)= " + export_key_Sitzungsnr + ") AND((RepSeriennr.Kennung) = \"Total Evaluation\" Or(RepSeriennr.Kennung) = \"Sr nr\" Or(RepSeriennr.Kennung) = \"Serial\" Or(RepSeriennr.Kennung) = \"Sr number\"));";

                OleDbDataReader reader = OleDbContext.ExecuteReader(query);

                RepSeriennrModel repSeriennrModel = new RepSeriennrModel();
                string ContainingColumns = "";
                repSeriennrModel.MPValStr = new List<string>();
                repSeriennrModel.MPColStr = new List<string>();

                while (reader.Read())
                {
                    repSeriennrModel.Kennung = reader["Kennung"].ToString();
                    
                    if (reader["Kennung"].ToString() != "Total Evaluation")
                    {
                        for (int i = 3; i < reader.FieldCount; i++)
                        {
                            string value = reader[i].ToString().Trim();
                            if (!string.IsNullOrEmpty(value))
                            {
                                //ContainingColumns = string.Join(", ", reader.GetName(i));
                                //repSeriennrModel.MPColStr.Add(reader.GetName(i));
                                repSeriennrModel.MPValStr.Add(value);                                                
                            }
                        }
                    }

                    if (reader["Kennung"].ToString() == "Total Evaluation")
                    {
                        for (int i = 3; i < reader.FieldCount; i++)
                        {
                            string value = reader[i].ToString().Trim();
                            if (!string.IsNullOrEmpty(value))
                            {
                                repSeriennrModel.MPColStr.Add(reader.GetName(i));
                                ContainingColumns = string.Join(", ", repSeriennrModel.MPColStr);
                            }
                        }
                    }

                }

                reader.Close();
                conn.Close();

                string SqlStr = "SELECT RepPP.PPName, RepPS.PSName, RepPSBez.Bezeichnung";

                //Condition for check if column available
                if (!string.IsNullOrEmpty(ContainingColumns.Trim()))
                    SqlStr += " , " + ContainingColumns;

                //Initialize full query
                SqlStr += " FROM ((RepPP INNER JOIN RepParam ON RepPP.[AutoID] = RepParam.[PPId]) INNER JOIN RepPS ON RepParam.[AutoID] = RepPS.[ParamId]) INNER JOIN RepPSBez ON RepPS.[AutoID] = RepPSBez.[RepPSId] WHERE(((RepPP.Sitzungsnr) = " + export_key_Sitzungsnr + ") AND((RepPSBez.Bezeichnung)Like\"Percentage%\")) ORDER BY RepPSBez.AutoID; ";

                conn.Open();
                OleDbCommand cmd = new OleDbCommand(SqlStr, conn);
                OleDbDataReader reader1 = cmd.ExecuteReader();
               
                Byte[] sheet = GetExcelDetail(reader1, repSeriennrModel);
                reader1.Close();
                conn.Close();
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "Save an Excel File";
                saveFileDialog.FileName = "Report_" + export_key_Sitzungsnr;

                //Show the SaveFileDialog and get the file path chosen by the user
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Write the byte array to a file
                    File.WriteAllBytes(saveFileDialog.FileName, sheet);
                    MessageBox.Show("Successfully exported excel file.");
                }

                dataGridReport.ColumnHeaderMouseClick += dataGridReport_ColumnHeaderMouseClick;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);              
            }
        }
        public Byte[] GetExcelDetail(OleDbDataReader reader, RepSeriennrModel resSernnModel)
        {

            var workbook = new XLWorkbook();
            MemoryStream fs = new MemoryStream();
            try
            {
                var workbookData = workbook.AddWorksheet();
                workbookData.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);

                workbookData.Cell("A1").Value = "PP Name";
                workbookData.Cell("A1").Style.Font.SetBold();

                workbookData.Cell("B1").Value = "PS Name";
                workbookData.Cell("B1").Style.Font.SetBold();

                workbookData.Cell("C1").Value = "Percentage Error";
                //workbookData.Cell("B1").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                workbookData.Cell("C1").Style.Font.SetBold(true);

                workbookData.Cell("D1").Value = "Meter Serial No.";
                //workbookData.Cell("C1").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                workbookData.Cell("D1").Style.Font.SetBold(true);

                workbookData.Cell("E1").Value = "Percentage Value";
                //workbookData.Cell("D1").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                workbookData.Cell("E1").Style.Font.SetBold(true);

                //string ppname = "";

                int rowindex = 2;
                while (reader.Read())
                {
                    //string cur_ppname = reader["PPName"].ToString();

                    /*if (cur_ppname != ppname)
                    {
                        //workbookData.Range("A" + rowindex.ToString() + ":D" + rowindex.ToString()).Merge(); // Merge cells for the current row
                        workbookData.Cell("A" + rowindex.ToString()).Value = cur_ppname;
                        workbookData.Cell("A" + rowindex.ToString()).Style.Font.SetBold();
                        rowindex += 1;
                        ppname = cur_ppname;
                    }*/
                    int innerRowCount = rowindex;
                    for (int i = 0; i < resSernnModel.MPColStr.Count; i++)
                    {
                        string columnName = resSernnModel.MPColStr[i].ToString();
                        if (!string.IsNullOrEmpty(reader[columnName].ToString()))
                        {
                            workbookData.Cell("A" + innerRowCount.ToString()).Value = reader["PPName"].ToString();

                            workbookData.Cell("B" + innerRowCount.ToString()).Value = reader["PSName"].ToString();
                            workbookData.Cell("C" + innerRowCount.ToString()).Value = reader["Bezeichnung"].ToString();
                            string meterSrNo = (i < resSernnModel.MPValStr.Count) ? resSernnModel.MPValStr[i] : "";
                            workbookData.Cell("D" + innerRowCount.ToString()).Value = meterSrNo;
                            workbookData.Cell("D" + innerRowCount.ToString()).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            workbookData.Cell("E" + innerRowCount.ToString()).Value = reader[columnName].ToString();
                            workbookData.Cell("E" + innerRowCount.ToString()).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            innerRowCount += 1;
                        }
                        else break;
                    }
                    rowindex = innerRowCount;
                }
                reader.Close();

                workbookData.Columns().AdjustToContents();
                workbook.SaveAs(fs);
                fs.Position = 0;
                return fs.ToArray();
            }

            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                workbook.Dispose();
                fs.Dispose();
            }

            return fs.ToArray();
        }


        private void dataGridReport_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // Make the cell uneditable
                dataGridReport[e.ColumnIndex, e.RowIndex].ReadOnly = true;
            }
        }

        private void btnChangeDatabase_Click(object sender, EventArgs e)
        {
           
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "mdb files (*.mdb)|*.mdb";
                openFileDialog.Title = "Select a mdb File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the selected file path from the OpenFileDialog
                    dbfilePath = openFileDialog.FileName;
                    string sql = "update FilePaths set mdbFilePath = '" + dbfilePath + "'";
                    int returnValue = DatabaseContext.ExecuteNonQuery(sql);
                    if (returnValue > 0)
                    {
                        frmReports_Load(sender, e);
                    }
                    // Reload the form with the new database file

                }
            return;
           
        }

        private void btnCancle_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void dataGridReport_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.Button == MouseButtons.Right)
            {
                // Assuming 'dataGridView1' is your DataGridView control
                dataGridReport.Rows[e.RowIndex].Selected = true;
                // Optionally, you can show a context menu or perform other actions here
                
            }
        }

        private void dataGridReport_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // A cell within the data area of the DataGridView was clicked

                // Make the cell uneditable
                dataGridReport.ReadOnly = true;

                try
                {
                    // Get the selected row
                    DataGridViewRow selectedRow = dataGridReport.Rows[e.RowIndex];

                    // Save the value of the "Sitzungsnr" cell in export_key_Sitzungsnr
                    export_key_Sitzungsnr = selectedRow.Cells["Sitzungsnr"].Value.ToString();

                    // Enable the button when a row is selected
                    btnExportToExcel.Enabled = true;
                }
                catch (Exception ex)
                {
                    // Handle the exception
                    MessageBox.Show("An error occurred: " + ex.Message);
                }
            }
            else if (e.RowIndex < 0 && e.ColumnIndex >= 0)
            {
                // A column header was clicked

                // Disable the button when a column header is clicked
                btnExportToExcel.Enabled = false;
            }
        }

        private void frmReports_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txtSearch.Text.Trim();

            if (!string.IsNullOrEmpty(searchTerm))
            {
                // Filter the DataGridView based on the search term
                DataView dv = ((DataTable)dataGridReport.DataSource).DefaultView;
                dv.RowFilter = $"PAName LIKE '%{searchTerm}%'"; // Use Contains for substring matching

                // Generate the serial numbers for the filtered rows
                for (int i = 0; i < dataGridReport.Rows.Count; i++)
                {
                    dataGridReport.Rows[i].Cells["SerialNumber"].Value = (i + 1).ToString();
                }
            }
            else
            {
                // If the search term is empty, clear the filter
                ((DataTable)dataGridReport.DataSource).DefaultView.RowFilter = "";

                // Re-generate serial numbers for all rows
                for (int i = 0; i < dataGridReport.Rows.Count; i++)
                {
                    dataGridReport.Rows[i].Cells["SerialNumber"].Value = (i + 1).ToString();
                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            // Clear the text in the search box
            txtSearch.Text = "";

            // Clear the filter applied to the DataGridView
            ((DataTable)dataGridReport.DataSource).DefaultView.RowFilter = "";

            // Re-populate the DataGridView with the original data
            frmReports_Load(sender, e);
        }

        private void dataGridReport_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

    }
}