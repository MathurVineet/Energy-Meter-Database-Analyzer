namespace DbExport
{
    partial class frmReports
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridReport = new System.Windows.Forms.DataGridView();
            this.btnExportToExcel = new System.Windows.Forms.Button();
            this.btnCancle = new System.Windows.Forms.Button();
            this.btnChangeDatabase = new System.Windows.Forms.Button();
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.btnClear = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridReport)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridReport
            // 
            this.dataGridReport.AllowUserToAddRows = false;
            this.dataGridReport.AllowUserToDeleteRows = false;
            this.dataGridReport.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dataGridReport.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridReport.Location = new System.Drawing.Point(12, 50);
            this.dataGridReport.Name = "dataGridReport";
            this.dataGridReport.ReadOnly = true;
            this.dataGridReport.RowHeadersVisible = false;
            this.dataGridReport.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridReport.Size = new System.Drawing.Size(1099, 657);
            this.dataGridReport.TabIndex = 0;
            this.dataGridReport.TabStop = false;
            this.dataGridReport.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridReport_CellClick);
            this.dataGridReport.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridReport_CellMouseClick);
            this.dataGridReport.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridReport_ColumnHeaderMouseClick);
            // 
            // btnExportToExcel
            // 
            this.btnExportToExcel.Enabled = false;
            this.btnExportToExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportToExcel.Location = new System.Drawing.Point(112, 713);
            this.btnExportToExcel.Name = "btnExportToExcel";
            this.btnExportToExcel.Size = new System.Drawing.Size(94, 32);
            this.btnExportToExcel.TabIndex = 4;
            this.btnExportToExcel.Text = "&Export";
            this.btnExportToExcel.UseVisualStyleBackColor = true;
            this.btnExportToExcel.Click += new System.EventHandler(this.btnExportToExcel_Click);
            // 
            // btnCancle
            // 
            this.btnCancle.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancle.Location = new System.Drawing.Point(1018, 713);
            this.btnCancle.Name = "btnCancle";
            this.btnCancle.Size = new System.Drawing.Size(94, 32);
            this.btnCancle.TabIndex = 5;
            this.btnCancle.Text = "&Close";
            this.btnCancle.UseVisualStyleBackColor = true;
            this.btnCancle.Click += new System.EventHandler(this.btnCancle_Click);
            // 
            // btnChangeDatabase
            // 
            this.btnChangeDatabase.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChangeDatabase.Location = new System.Drawing.Point(12, 713);
            this.btnChangeDatabase.Name = "btnChangeDatabase";
            this.btnChangeDatabase.Size = new System.Drawing.Size(94, 32);
            this.btnChangeDatabase.TabIndex = 3;
            this.btnChangeDatabase.Text = "&Select DB";
            this.btnChangeDatabase.UseVisualStyleBackColor = true;
            this.btnChangeDatabase.Click += new System.EventHandler(this.btnChangeDatabase_Click);
            // 
            // txtSearch
            // 
            this.txtSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSearch.Location = new System.Drawing.Point(12, 12);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(1030, 30);
            this.txtSearch.TabIndex = 1;
            this.txtSearch.TextChanged += new System.EventHandler(this.txtSearch_TextChanged);
            // 
            // btnClear
            // 
            this.btnClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClear.Location = new System.Drawing.Point(1052, 12);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(59, 32);
            this.btnClear.TabIndex = 2;
            this.btnClear.Text = "&Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // frmReports
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1124, 751);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.txtSearch);
            this.Controls.Add(this.btnChangeDatabase);
            this.Controls.Add(this.btnCancle);
            this.Controls.Add(this.btnExportToExcel);
            this.Controls.Add(this.dataGridReport);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "frmReports";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Report";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmReports_FormClosing);
            this.Load += new System.EventHandler(this.frmReports_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridReport)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridReport;
        private System.Windows.Forms.Button btnExportToExcel;
        private System.Windows.Forms.Button btnCancle;
        private System.Windows.Forms.Button btnChangeDatabase;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.Button btnClear;
    }
}

