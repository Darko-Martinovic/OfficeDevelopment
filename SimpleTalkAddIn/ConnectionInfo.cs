using System;
using System.Data;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Core;
using SimpleTalkExcellAddin.Utils;
using ListBox = System.Windows.Forms.ListBox;


namespace SimpleTalkExcellAddin
{
    public partial class ConnectionInfo : Form
    {

        public ConnectionInfo()
        {
            InitializeComponent();
        }


        private void ConnectionInfo_Load(object sender, EventArgs e)
        {

            mtbMain.SelectedTab = tbConnection;
            cmbAuth.SelectedIndex = 0;

            cmbCharTypes.DataSource = Enum.GetValues(typeof(XlChartType));
            cmbCharTypes.SelectedItem = XlChartType.xl3DColumnClustered;

            // There are 21 table Light styles in Excel 2013(6), 28 Medium styles and 11 dark styles
            for (var i = 1; i < 22; i++)
            {
                cmbTableStyles.Items.Add($"TableStyleLight{i.ToString()}");
                cmbPivotStyles.Items.Add($"PivotStyleLight{i.ToString()}");
            }
            for (var i = 1; i < 28; i++)
            {
                cmbTableStyles.Items.Add($"TableStyleMedium{i.ToString()}");
                cmbPivotStyles.Items.Add($"PivotStyleMedium{i.ToString()}");
            }
            for (var i = 1; i < 11; i++)
            {
                cmbTableStyles.Items.Add($"TableStyleDark{i.ToString()}");
                cmbPivotStyles.Items.Add($"PivotStyleDark{i.ToString()}");
            }

            cmbTableStyles.SelectedItem = @"TableStyleLight7";
            cmbPivotStyles.SelectedItem = @"PivotStyleLight16";

            txtServer.Focus();
            if (!Debugger.IsAttached) return;

            txtQuery.Text = @"

SELECT [SalesOrderID]
      ,[RevisionNumber]
      ,[OrderDate]
      ,[DueDate]
      ,[ShipDate]
      ,[Status]
      ,[OnlineOrderFlag]
      ,[SalesOrderNumber]
      ,[PurchaseOrderNumber]
      ,[AccountNumber]
      ,[CustomerID]
      ,[SalesPersonID]
      ,[TerritoryID]
      ,[BillToAddressID]
      ,[ShipToAddressID]
      ,[ShipMethodID]
      ,[CreditCardID]
      ,[CreditCardApprovalCode]
      ,[CurrencyRateID]
      ,[SubTotal]
      ,[TaxAmt]
      ,[Freight]
      ,[TotalDue]
      ,[Comment]

  FROM [Sales].[SalesOrderHeader]
";

            cmbDatabase.Text = @"AdventureWorks2016";
        }

        private void cmbAuth_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbAuth.SelectedItem != null)
            {
                if (_isError)
                    _isError = false;

                cmbDatabase.SelectedIndex = -1;

                if (cmbAuth.SelectedIndex == 0)
                {
                    txtUserName.Text = "";
                    txtPassword.Text = "";
                    txtUserName.Enabled = false;
                    txtPassword.Enabled = false;
                }
                else if (cmbAuth.SelectedIndex == 1)
                {
                    txtUserName.Enabled = true;
                    txtPassword.Enabled = true;
                    txtUserName.Select();
                }
            }
        }

        private bool _isError;


        private void cmbDatabase_MouseClick(object sender, MouseEventArgs e)
        {
            if (txtServer.Text.Trim().Equals(string.Empty))
            {
                txtServer.Focus();
                MessageBox.Show(@"Please enter the valid server name",
                    @"Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;

            }

            if (cmbAuth.SelectedIndex != 0 && (txtUserName.Text.Trim().Equals(string.Empty) ||
                                               txtPassword.Text.Trim().Equals(string.Empty)))
            {
                MessageBox.Show(@"Please enter userName and password",
                    @"Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (_isError)
                    _isError = false;
                return;
            }

            if (_isError == false)
            {
                BindDataBases(cmbDatabase);
            }
        }

        private void BindDataBases(ComboBox cmb)
        {
            cmb.Items.Clear();
            var ds = DataAccess.GetDataSet(
                DataAccess.GetConnectionString(
                    txtServer.Text,
                    "master",
                    cmbAuth.SelectedIndex == 0,
                    txtUserName.Text, txtPassword.Text), @"SELECT name 
                                                                FROM sys.databases
                                                                WHERE state = 0 
                                                                    AND is_read_only = 0 
                                                                ORDER BY name", null, out var error);
            if (error.Equals(string.Empty) == false)
            {
                _isError = true;
                MessageBox.Show($@"Error binding database information : {error}",
                    @"Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                // ReSharper disable once RedundantAssignment
                ds = null;
            }
            else
            {
                foreach (DataRow r in ds.Tables[0].Rows)
                    cmb.Items.Add(r["Name"].ToString());
                // ReSharper disable once RedundantAssignment
                ds = null;

            }

        }

        public Inputs MyInputs { get; private set; }

        private void mtbExecute_Click(object sender, EventArgs e)
        {
            if (txtServer.Text.Equals(string.Empty))
            {
                txtServer.Focus();
                MessageBox.Show(@"Please enter the valid server name",
                    @"Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (cmbAuth.SelectedIndex == -1)
            {
                cmbAuth.Focus();
                MessageBox.Show(@"Please choose authentication type!",
                    @"Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (cmbDatabase.SelectedIndex == -1)
            {
                cmbDatabase.Focus();
                MessageBox.Show(@"Please choose database",
                    @"Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (lstMain.Items.Count == 0)
            {
                MessageBox.Show(@"Please click on 'Refresh' button",
                    @"Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                mtbMain.SelectedTab = tbFields;
                return;

            }

            MyInputs = new Inputs
            {

                Query = txtQuery.Text,
                ConnectionString = Inputs.GetConnectionString(txtServer.Text, cmbDatabase.Text,
                    cmbAuth.SelectedIndex == 0, txtUserName.Text, txtPassword.Text),
                ChartType = (XlChartType) cmbCharTypes.SelectedItem,
                TableStyle = cmbTableStyles.SelectedItem.ToString(),
                PivotStyle = cmbPivotStyles.SelectedItem.ToString(),
            };


            foreach (string s in lstMain.Items)
                MyInputs.AllFields.Add(s);

            foreach (string s in lstReportFilters.Items)
                MyInputs.ReportFielters.Add(s);

            foreach (string s in lstColumns.Items)
                MyInputs.Columns.Add(s);

            foreach (string s in lstRows.Items)
                MyInputs.Rows.Add(s);

            foreach (string s in lstValues.Items)
                MyInputs.Values.Add(s);

            DialogResult = DialogResult.OK;
            Close();

        }

        private void LstMain_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.StringFormat))
            {
                string str = (string)e.Data.GetData(DataFormats.StringFormat);
                ((ListBox)sender).Items.Add(str);
            }
        }

        private void LstMain_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void LstMain_MouseDown(object sender, MouseEventArgs e)
        {

            var lb = (ListBox)sender;

            if (lb.Items.Count == 0)
                return;
            string s = lb.Items[lb.IndexFromPoint(e.X, e.Y)].ToString();
            DragDropEffects dde1 = DoDragDrop(s,
                DragDropEffects.All);

            if (dde1 == DragDropEffects.All)
            {
                lb.Items.RemoveAt(lb.IndexFromPoint(e.X, e.Y));
            }

        }

        private void mtbMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (mtbMain.SelectedTab == tbFields && lstMain.Items.Count == 0)
            {
                if (txtServer.Text.Trim().Equals(string.Empty))
                {
                    txtServer.Focus();
                    MessageBox.Show(@"Please enter the valid server name",
                        @"Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;

                }
                if (cmbAuth.SelectedIndex != 0 && (txtUserName.Text.Trim().Equals(string.Empty) ||
                                                   txtPassword.Text.Trim().Equals(string.Empty)))
                {
                    MessageBox.Show(@"Please enter userName and password",
                        @"Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (_isError)
                        _isError = false;
                    return;
                }

                if (cmbDatabase.SelectedIndex == 0 || cmbDatabase.Text.Equals(string.Empty))
                {
                    MessageBox.Show(@"Please select the database",
                        @"Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (_isError)
                        _isError = false;
                    return;
                }


                var command = $" SET FMTONLY ON {Environment.NewLine}{txtQuery.Text}{Environment.NewLine}SET FMTONLY OFF";

                var ds = DataAccess.GetDataSet(
                    DataAccess.GetConnectionString(
                        txtServer.Text,
                        cmbDatabase.Text,
                        cmbAuth.SelectedIndex == 0,
                        txtUserName.Text, txtPassword.Text), command, null, out var error);
                if (error.Equals(string.Empty) == false)
                {
                    _isError = true;
                    MessageBox.Show($@"Error binding database information : {error}",
                        @"Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    // ReSharper disable once RedundantAssignment
                    ds = null;
                }
                else
                {
                    lstMain.Items.Clear();
                    lstColumns.Items.Clear();
                    lstRows.Items.Clear();
                    lstReportFilters.Items.Clear();
                    lstValues.Items.Clear();

                    foreach (DataColumn c in ds.Tables[0].Columns)
                    {
                        //if (c.DataType == Type.GetType("System.DateTime") && lstColumns.Items.Count == 0)
                        //{
                        //    lstColumns.Items.Add(c.ColumnName);
                        //}
                        //else if (c.DataType == Type.GetType("System.Decimal") && lstValues.Items.Count == 0)
                        //{
                        //    lstValues.Items.Add(c.ColumnName);
                        //}
                        //else if ((c.DataType == Type.GetType("System.String") || c.DataType == Type.GetType("System.Int32")) && lstRows.Items.Count == 0)
                        //{
                        //    lstRows.Items.Add(c.ColumnName);
                        //}
                        //else if (c.DataType == Type.GetType("System.DateTime") && lstReportFilters.Items.Count == 0)
                        //{
                        //    lstReportFilters.Items.Add(c.ColumnName);
                        //}
                        //else
                        //{
                        lstMain.Items.Add(c.ColumnName);
                        //}
                    }

                    if (Debugger.IsAttached)
                    {


                        if (lstMain.Items.Contains("SalesPersonID"))
                        {
                            lstMain.Items.Remove("SalesPersonID");
                            lstColumns.Items.Add("SalesPersonID");
                        }
                        if (lstMain.Items.Contains("TerritoryID"))
                        {
                            lstMain.Items.Remove("TerritoryID");
                            lstRows.Items.Add("TerritoryID");
                        }
                        if (lstMain.Items.Contains("TotalDue"))
                        {
                            lstMain.Items.Remove("TotalDue");
                            lstValues.Items.Add("TotalDue");
                        }
                        if (lstMain.Items.Contains("RevisionNumber"))
                        {
                            lstMain.Items.Remove("RevisionNumber");
                            lstReportFilters.Items.Add("RevisionNumber");
                        }



                    }
                    // ReSharper disable once RedundantAssignment
                    ds = null;

                }

            }
        }
    }
    //});

}


