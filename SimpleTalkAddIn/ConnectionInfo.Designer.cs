using System.Windows.Forms;

namespace SimpleTalkExcellAddin
{
    partial class ConnectionInfo
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
            this.mtbMain = new System.Windows.Forms.TabControl();
            this.tbConnection = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.txtServer = new System.Windows.Forms.TextBox();
            this.cmbDatabase = new System.Windows.Forms.ComboBox();
            this.cmbAuth = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.tbQuery = new System.Windows.Forms.TabPage();
            this.metroLabel1 = new System.Windows.Forms.Label();
            this.txtQuery = new System.Windows.Forms.TextBox();
            this.tbFields = new System.Windows.Forms.TabPage();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lstRows = new System.Windows.Forms.ListBox();
            this.lstValues = new System.Windows.Forms.ListBox();
            this.lstColumns = new System.Windows.Forms.ListBox();
            this.lstReportFilters = new System.Windows.Forms.ListBox();
            this.lstMain = new System.Windows.Forms.ListBox();
            this.tbOptions = new System.Windows.Forms.TabPage();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.cmbPivotStyles = new System.Windows.Forms.ComboBox();
            this.cmbTableStyles = new System.Windows.Forms.ComboBox();
            this.cmbCharTypes = new System.Windows.Forms.ComboBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.mtbMain.SuspendLayout();
            this.tbConnection.SuspendLayout();
            this.tbQuery.SuspendLayout();
            this.tbFields.SuspendLayout();
            this.tbOptions.SuspendLayout();
            this.SuspendLayout();
            // 
            // mtbMain
            // 
            this.mtbMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.mtbMain.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.mtbMain.Controls.Add(this.tbConnection);
            this.mtbMain.Controls.Add(this.tbQuery);
            this.mtbMain.Controls.Add(this.tbFields);
            this.mtbMain.Controls.Add(this.tbOptions);
            this.mtbMain.Location = new System.Drawing.Point(0, 0);
            this.mtbMain.Name = "mtbMain";
            this.mtbMain.SelectedIndex = 0;
            this.mtbMain.Size = new System.Drawing.Size(646, 442);
            this.mtbMain.TabIndex = 58;
            this.mtbMain.SelectedIndexChanged += new System.EventHandler(this.mtbMain_SelectedIndexChanged);
            // 
            // tbConnection
            // 
            this.tbConnection.Controls.Add(this.label1);
            this.tbConnection.Controls.Add(this.txtServer);
            this.tbConnection.Controls.Add(this.cmbDatabase);
            this.tbConnection.Controls.Add(this.cmbAuth);
            this.tbConnection.Controls.Add(this.label5);
            this.tbConnection.Controls.Add(this.label4);
            this.tbConnection.Controls.Add(this.label3);
            this.tbConnection.Controls.Add(this.label2);
            this.tbConnection.Controls.Add(this.txtPassword);
            this.tbConnection.Controls.Add(this.txtUserName);
            this.tbConnection.Location = new System.Drawing.Point(4, 25);
            this.tbConnection.Name = "tbConnection";
            this.tbConnection.Size = new System.Drawing.Size(638, 413);
            this.tbConnection.TabIndex = 0;
            this.tbConnection.Text = "Connection";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label1.Location = new System.Drawing.Point(37, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 15);
            this.label1.TabIndex = 70;
            this.label1.Text = "Server name :";
            // 
            // txtServer
            // 
            this.txtServer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtServer.Location = new System.Drawing.Point(37, 44);
            this.txtServer.Name = "txtServer";
            this.txtServer.Size = new System.Drawing.Size(247, 20);
            this.txtServer.TabIndex = 71;
            // 
            // cmbDatabase
            // 
            this.cmbDatabase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDatabase.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmbDatabase.FormattingEnabled = true;
            this.cmbDatabase.ItemHeight = 13;
            this.cmbDatabase.Location = new System.Drawing.Point(39, 296);
            this.cmbDatabase.Name = "cmbDatabase";
            this.cmbDatabase.Size = new System.Drawing.Size(249, 21);
            this.cmbDatabase.TabIndex = 68;
            this.cmbDatabase.MouseClick += new System.Windows.Forms.MouseEventHandler(this.cmbDatabase_MouseClick);
            // 
            // cmbAuth
            // 
            this.cmbAuth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAuth.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmbAuth.FormattingEnabled = true;
            this.cmbAuth.ItemHeight = 13;
            this.cmbAuth.Items.AddRange(new object[] {
            "Windows authentication",
            "Sql server authentication"});
            this.cmbAuth.Location = new System.Drawing.Point(39, 100);
            this.cmbAuth.Name = "cmbAuth";
            this.cmbAuth.Size = new System.Drawing.Size(249, 21);
            this.cmbAuth.TabIndex = 65;
            this.cmbAuth.SelectedIndexChanged += new System.EventHandler(this.cmbAuth_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label5.Location = new System.Drawing.Point(39, 274);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(104, 15);
            this.label5.TabIndex = 58;
            this.label5.Text = "Database name  :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label4.Location = new System.Drawing.Point(39, 212);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(67, 15);
            this.label4.TabIndex = 59;
            this.label4.Text = "Password :";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label3.Location = new System.Drawing.Point(37, 141);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 15);
            this.label3.TabIndex = 60;
            this.label3.Text = "User name  :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label2.Location = new System.Drawing.Point(37, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(90, 15);
            this.label2.TabIndex = 61;
            this.label2.Text = "Authentication :";
            // 
            // txtPassword
            // 
            this.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtPassword.Enabled = false;
            this.txtPassword.Location = new System.Drawing.Point(39, 234);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(247, 20);
            this.txtPassword.TabIndex = 67;
            // 
            // txtUserName
            // 
            this.txtUserName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtUserName.Enabled = false;
            this.txtUserName.Location = new System.Drawing.Point(39, 163);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(247, 20);
            this.txtUserName.TabIndex = 66;
            // 
            // tbQuery
            // 
            this.tbQuery.Controls.Add(this.metroLabel1);
            this.tbQuery.Controls.Add(this.txtQuery);
            this.tbQuery.Location = new System.Drawing.Point(4, 25);
            this.tbQuery.Name = "tbQuery";
            this.tbQuery.Size = new System.Drawing.Size(638, 413);
            this.tbQuery.TabIndex = 1;
            this.tbQuery.Text = "Query";
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.metroLabel1.Location = new System.Drawing.Point(3, 15);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(45, 15);
            this.metroLabel1.TabIndex = 71;
            this.metroLabel1.Text = "Query :";
            // 
            // txtQuery
            // 
            this.txtQuery.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtQuery.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtQuery.Location = new System.Drawing.Point(3, 49);
            this.txtQuery.Multiline = true;
            this.txtQuery.Name = "txtQuery";
            this.txtQuery.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtQuery.Size = new System.Drawing.Size(632, 322);
            this.txtQuery.TabIndex = 2;
            // 
            // tbFields
            // 
            this.tbFields.Controls.Add(this.label9);
            this.tbFields.Controls.Add(this.label10);
            this.tbFields.Controls.Add(this.label8);
            this.tbFields.Controls.Add(this.label7);
            this.tbFields.Controls.Add(this.label6);
            this.tbFields.Controls.Add(this.lstRows);
            this.tbFields.Controls.Add(this.lstValues);
            this.tbFields.Controls.Add(this.lstColumns);
            this.tbFields.Controls.Add(this.lstReportFilters);
            this.tbFields.Controls.Add(this.lstMain);
            this.tbFields.Location = new System.Drawing.Point(4, 25);
            this.tbFields.Name = "tbFields";
            this.tbFields.Size = new System.Drawing.Size(638, 413);
            this.tbFields.TabIndex = 2;
            this.tbFields.Text = "Fields";
            this.tbFields.MouseDown += new System.Windows.Forms.MouseEventHandler(this.LstMain_MouseDown);
            // 
            // label9
            // 
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label9.AutoSize = true;
            this.label9.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label9.Location = new System.Drawing.Point(451, 280);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(44, 15);
            this.label9.TabIndex = 72;
            this.label9.Text = "ROWS";
            // 
            // label10
            // 
            this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label10.AutoSize = true;
            this.label10.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label10.Location = new System.Drawing.Point(227, 145);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(53, 15);
            this.label10.TabIndex = 72;
            this.label10.Text = "VALUES";
            // 
            // label8
            // 
            this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label8.AutoSize = true;
            this.label8.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label8.Location = new System.Drawing.Point(451, 145);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(68, 15);
            this.label8.TabIndex = 72;
            this.label8.Text = "COLUMNS";
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label7.AutoSize = true;
            this.label7.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label7.Location = new System.Drawing.Point(451, 22);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(56, 15);
            this.label7.TabIndex = 72;
            this.label7.Text = "FILTERS";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label6.Location = new System.Drawing.Point(8, 16);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(172, 15);
            this.label6.TabIndex = 72;
            this.label6.Text = "Choose fields to add to report :";
            // 
            // lstRows
            // 
            this.lstRows.AllowDrop = true;
            this.lstRows.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lstRows.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lstRows.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstRows.HorizontalScrollbar = true;
            this.lstRows.ItemHeight = 17;
            this.lstRows.Location = new System.Drawing.Point(454, 299);
            this.lstRows.Name = "lstRows";
            this.lstRows.Size = new System.Drawing.Size(176, 53);
            this.lstRows.TabIndex = 1;
            this.lstRows.DragDrop += new System.Windows.Forms.DragEventHandler(this.LstMain_DragDrop);
            this.lstRows.DragOver += new System.Windows.Forms.DragEventHandler(this.LstMain_DragOver);
            this.lstRows.MouseDown += new System.Windows.Forms.MouseEventHandler(this.LstMain_MouseDown);
            // 
            // lstValues
            // 
            this.lstValues.AllowDrop = true;
            this.lstValues.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lstValues.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lstValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstValues.HorizontalScrollbar = true;
            this.lstValues.ItemHeight = 17;
            this.lstValues.Location = new System.Drawing.Point(230, 164);
            this.lstValues.Name = "lstValues";
            this.lstValues.Size = new System.Drawing.Size(176, 53);
            this.lstValues.TabIndex = 1;
            this.lstValues.DragDrop += new System.Windows.Forms.DragEventHandler(this.LstMain_DragDrop);
            this.lstValues.DragOver += new System.Windows.Forms.DragEventHandler(this.LstMain_DragOver);
            this.lstValues.MouseDown += new System.Windows.Forms.MouseEventHandler(this.LstMain_MouseDown);
            // 
            // lstColumns
            // 
            this.lstColumns.AllowDrop = true;
            this.lstColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lstColumns.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lstColumns.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstColumns.HorizontalScrollbar = true;
            this.lstColumns.ItemHeight = 17;
            this.lstColumns.Location = new System.Drawing.Point(454, 164);
            this.lstColumns.Name = "lstColumns";
            this.lstColumns.Size = new System.Drawing.Size(176, 53);
            this.lstColumns.TabIndex = 1;
            this.lstColumns.DragDrop += new System.Windows.Forms.DragEventHandler(this.LstMain_DragDrop);
            this.lstColumns.DragOver += new System.Windows.Forms.DragEventHandler(this.LstMain_DragOver);
            this.lstColumns.MouseDown += new System.Windows.Forms.MouseEventHandler(this.LstMain_MouseDown);
            // 
            // lstReportFilters
            // 
            this.lstReportFilters.AllowDrop = true;
            this.lstReportFilters.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lstReportFilters.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lstReportFilters.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstReportFilters.HorizontalScrollbar = true;
            this.lstReportFilters.ItemHeight = 17;
            this.lstReportFilters.Location = new System.Drawing.Point(454, 41);
            this.lstReportFilters.Name = "lstReportFilters";
            this.lstReportFilters.Size = new System.Drawing.Size(176, 53);
            this.lstReportFilters.TabIndex = 1;
            this.lstReportFilters.DragDrop += new System.Windows.Forms.DragEventHandler(this.LstMain_DragDrop);
            this.lstReportFilters.DragOver += new System.Windows.Forms.DragEventHandler(this.LstMain_DragOver);
            this.lstReportFilters.MouseDown += new System.Windows.Forms.MouseEventHandler(this.LstMain_MouseDown);
            // 
            // lstMain
            // 
            this.lstMain.AllowDrop = true;
            this.lstMain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lstMain.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstMain.HorizontalScrollbar = true;
            this.lstMain.ItemHeight = 17;
            this.lstMain.Location = new System.Drawing.Point(11, 41);
            this.lstMain.Name = "lstMain";
            this.lstMain.Size = new System.Drawing.Size(176, 359);
            this.lstMain.TabIndex = 1;
            this.lstMain.DragDrop += new System.Windows.Forms.DragEventHandler(this.LstMain_DragDrop);
            this.lstMain.DragOver += new System.Windows.Forms.DragEventHandler(this.LstMain_DragOver);
            this.lstMain.MouseDown += new System.Windows.Forms.MouseEventHandler(this.LstMain_MouseDown);
            // 
            // tbOptions
            // 
            this.tbOptions.Controls.Add(this.label13);
            this.tbOptions.Controls.Add(this.label12);
            this.tbOptions.Controls.Add(this.label11);
            this.tbOptions.Controls.Add(this.cmbPivotStyles);
            this.tbOptions.Controls.Add(this.cmbTableStyles);
            this.tbOptions.Controls.Add(this.cmbCharTypes);
            this.tbOptions.Location = new System.Drawing.Point(4, 25);
            this.tbOptions.Name = "tbOptions";
            this.tbOptions.Size = new System.Drawing.Size(638, 413);
            this.tbOptions.TabIndex = 3;
            this.tbOptions.Text = "Options";
            this.tbOptions.UseVisualStyleBackColor = true;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(20, 91);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(66, 15);
            this.label13.TabIndex = 1;
            this.label13.Text = "Pivot styles";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(20, 54);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(73, 15);
            this.label12.TabIndex = 1;
            this.label12.Text = "Table Styles";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(20, 20);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(71, 15);
            this.label11.TabIndex = 1;
            this.label11.Text = "Chart Types";
            // 
            // cmbPivotStyles
            // 
            this.cmbPivotStyles.FormattingEnabled = true;
            this.cmbPivotStyles.Location = new System.Drawing.Point(98, 88);
            this.cmbPivotStyles.Name = "cmbPivotStyles";
            this.cmbPivotStyles.Size = new System.Drawing.Size(202, 21);
            this.cmbPivotStyles.TabIndex = 0;
            // 
            // cmbTableStyles
            // 
            this.cmbTableStyles.FormattingEnabled = true;
            this.cmbTableStyles.Location = new System.Drawing.Point(98, 51);
            this.cmbTableStyles.Name = "cmbTableStyles";
            this.cmbTableStyles.Size = new System.Drawing.Size(202, 21);
            this.cmbTableStyles.TabIndex = 0;
            // 
            // cmbCharTypes
            // 
            this.cmbCharTypes.FormattingEnabled = true;
            this.cmbCharTypes.Location = new System.Drawing.Point(98, 17);
            this.cmbCharTypes.Name = "cmbCharTypes";
            this.cmbCharTypes.Size = new System.Drawing.Size(202, 21);
            this.cmbCharTypes.TabIndex = 0;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(393, 463);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(249, 23);
            this.btnCancel.TabIndex = 62;
            this.btnCancel.TabStop = false;
            this.btnCancel.Text = "Cancel";
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(12, 463);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(249, 23);
            this.btnGenerate.TabIndex = 75;
            this.btnGenerate.TabStop = false;
            this.btnGenerate.Text = "Execute";
            this.btnGenerate.Click += new System.EventHandler(this.mtbExecute_Click);
            // 
            // ConnectionInfo
            // 
            this.AcceptButton = this.btnGenerate;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(646, 498);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.mtbMain);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ConnectionInfo";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Simple Talk Excel Addin";
            this.Load += new System.EventHandler(this.ConnectionInfo_Load);
            this.mtbMain.ResumeLayout(false);
            this.tbConnection.ResumeLayout(false);
            this.tbConnection.PerformLayout();
            this.tbQuery.ResumeLayout(false);
            this.tbQuery.PerformLayout();
            this.tbFields.ResumeLayout(false);
            this.tbFields.PerformLayout();
            this.tbOptions.ResumeLayout(false);
            this.tbOptions.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TabControl mtbMain;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtServer;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ComboBox cmbDatabase;
        private System.Windows.Forms.ComboBox cmbAuth;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.Label metroLabel1;
        private System.Windows.Forms.TextBox txtQuery;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.TabPage tbConnection;
        private System.Windows.Forms.TabPage tbQuery;
        private System.Windows.Forms.TabPage tbFields;
        private Label label9;
        private Label label10;
        private Label label8;
        private Label label7;
        private Label label6;
        private ListBox lstRows;
        private ListBox lstValues;
        private ListBox lstColumns;
        private ListBox lstReportFilters;
        private ListBox lstMain;
        private TabPage tbOptions;
        private ComboBox cmbCharTypes;
        private Label label11;
        private Label label12;
        private ComboBox cmbTableStyles;
        private Label label13;
        private ComboBox cmbPivotStyles;
    }
}