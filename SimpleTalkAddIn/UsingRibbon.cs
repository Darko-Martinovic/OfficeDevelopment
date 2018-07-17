using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Cthread = System.Threading.Thread;
using SimpleTalkExcellAddin.Utils;

    //How to make pfx
    //MakeCert /n "CN=SimpleTalk" /r /h 0 /eku "1.3.6.1.5.5.7.3.3,1.3.6.1.4.1.311.10.3.13" /e "01/16/2174" /sv SimpleTalk.pvk SimpleTalk.cer /a 
    //pvk2pfx -pvk SimpleTalk.pvk -spc SimpleTalk.cer -pfx SimpleTalk.pfx –f
    //--


namespace SimpleTalkExcellAddin
{
    public partial class UsingRibbon
    {
        private Inputs _myInput;

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var inputForm = new ConnectionInfo();

            if (inputForm.ShowDialog() == DialogResult.OK)
            {
                _myInput = inputForm.MyInputs;

            }

            if (_myInput == null)
                return;

            var eWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            var sheets = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets;
            var exists = false;

            var counter = 0;
            foreach (var mSheet in sheets)
            {
                if (!(mSheet is Excel.Worksheet t1)) continue;

                if (!t1.Name.StartsWith("Simple Talk")) continue;

                exists = true;
                counter++;
            }
            if (exists && MessageBox.Show(@"You have already created a set of tables. Do you want to add new ones?", @"Question",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;


            var t = Task.Run(() => Helper.GetSource(_myInput));


            t.ContinueWith(a => MakeExcel(pivotStyle: _myInput.PivotStyle,
                tableStyle: _myInput.TableStyle,
                charType: _myInput.ChartType,
                dataSource: t.Result,
                eWorkbook: eWorkbook,
                counter: counter,
                sheetTitle: @"Simple Talk",
                pivotTableName: @"Pivot table Simple Talk",
                pivotRowFields: _myInput.Rows,
                pivotValueFields: _myInput.Values,
                pivotColumnFields: _myInput.Columns,
                pivotReportFields: _myInput.ReportFielters,
                sliderFields: _myInput.ReportFielters));
            t.Wait();

        }




        private void btnNoInput_Click(object sender, RibbonControlEventArgs e)
        {

            //names.Split(',').ToList();
            _myInput = new Inputs
            {
                ConnectionString = ConfigurationManager.AppSettings["ConnStr"],
                Query = ConfigurationManager.AppSettings["Query"],
                Values = ConfigurationManager.AppSettings["Values"].Split(',').ToList(),
                Rows = ConfigurationManager.AppSettings["Rows"].Split(',').ToList(),
                Columns = ConfigurationManager.AppSettings["Columns"].Split(',').ToList(),
                ReportFielters = ConfigurationManager.AppSettings["ReportFielters"].Split(',').ToList(),
                TableStyle = ConfigurationManager.AppSettings["TableStyle"],
                PivotStyle = ConfigurationManager.AppSettings["PivotStyle"],
                ChartType = (XlChartType)Enum.Parse(typeof(XlChartType), ConfigurationManager.AppSettings["ChartType"])
        };



            var eWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            var sheets = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets;
            var exists = false;

            var counter = 0;
            foreach (var mSheet in sheets)
            {
                if (!(mSheet is Excel.Worksheet t1)) continue;

                if (!t1.Name.StartsWith("Simple Talk")) continue;

                exists = true;
                counter++;
            }
            if (exists && MessageBox.Show(@"You have already created a set of tables. Do you want to add new ones?", @"Question",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;


            var t = Task.Run(() => Helper.GetSource(_myInput));


            t.ContinueWith(a => MakeExcel(pivotStyle: _myInput.PivotStyle,
                tableStyle: _myInput.TableStyle,
                charType: _myInput.ChartType,
                dataSource: t.Result,
                eWorkbook: eWorkbook,
                counter: counter,
                sheetTitle: @"Simple Talk",
                pivotTableName: @"Pivot table Simple Talk",
                pivotRowFields: _myInput.Rows,
                pivotValueFields: _myInput.Values,
                pivotColumnFields: _myInput.Columns,
                pivotReportFields: _myInput.ReportFielters,
                sliderFields: _myInput.ReportFielters));
            t.Wait();


        }


        private static void MakeExcel(
          string pivotStyle,
          string tableStyle,
          XlChartType charType,
          DataTable dataSource,
          Excel.Workbook eWorkbook,
          int counter,
          string sheetTitle,
          string pivotTableName,
          List<string> pivotRowFields,
          List<string> pivotValueFields,
          List<string> pivotColumnFields,
          List<string> pivotReportFields,
          List<string> sliderFields)
        {
            var startRow = 1;
            var endIndex = dataSource.Rows.Count + startRow;

            var sheets = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets;
            var sheet = (Excel.Worksheet)sheets.Add();

            sheet.Name = sheetTitle + (++counter);

            for (var i = 1; i < dataSource.Columns.Count + 1; i++)
            {
                sheet.Cells[startRow, i] = dataSource.Columns[i - 1].ColumnName;
                if (dataSource.Columns[i - 1].DataType == Type.GetType("System.DateTime"))
                {
                    sheet.Range[sheet.Cells[startRow + 1, i], sheet.Cells[endIndex, i]].NumberFormat =
                        Cthread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern;

                }
                else if (dataSource.Columns[i - 1].DataType == Type.GetType("System.Decimal"))
                {
                    sheet.Range[sheet.Cells[startRow + 1, i], sheet.Cells[endIndex, i]].NumberFormat =
                        $"#{Cthread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator}##0{Cthread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator}00";

                }

            }

            try
            {
                sheet.Range[sheet.Cells[startRow + 1, 1], sheet.Cells[endIndex, dataSource.Columns.Count]].Value2 =
                    General.Convert(dataSource);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

            General.FormatAsTable(sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[endIndex, dataSource.Columns.Count]],
                "SourceData" + counter, tableStyle);


            //eApplication.Visible = true;
            //sheet.Activate();
            //sheet.Application.ActiveWindow.SplitRow = 1;
            //sheet.Application.ActiveWindow.SplitColumn = 1;
            //sheet.Application.ActiveWindow.FreezePanes = true;
            //// Now apply autofilter
            //var firstRow = (Excel.Range)sheet.Rows[1];

            //firstRow.AutoFilter(1,
            //    Type.Missing,
            //    Excel.XlAutoFilterOperator.xlAnd,
            //    Type.Missing,
            //    true);

            sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[endIndex, dataSource.Columns.Count]].Columns.AutoFit();
            //sheet.Columns.AutoFit();


            var pivotData = sheet.Range["A1",
                General.GetColumnName(dataSource.Columns.Count) + (dataSource.Rows.Count + 1)];



            var pivotWorkSheet = General.AddPivot(
                pivotStyle,
                pivotData,
                dataSource,
                counter,
                pivotTableName,
                pivotRowFields,
                pivotColumnFields,
                pivotValueFields,
                pivotReportFields);





            General.AddChart(pivotWorkSheet, sheetTitle, pivotData, charType);

            General.AddSlicers(eWorkbook, (Excel.PivotTable)pivotWorkSheet.PivotTables(pivotTableName),
                pivotWorkSheet, sliderFields);

            pivotWorkSheet.Select();

            //Excel.Application oApp = Globals.ThisAddIn.Application;
            //oApp.Visible = true;
            //oApp.ScreenUpdating = true;
            //oApp.UserControl = true;
            //oApp.Interactive = true;


        }

    }


}
