using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using ExportToExcel.Data;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace ExportToExcel.Excel
{
    public static class ExportToExcel
    {
        public static void Export(bool pasteRange, 
                                  DataSet ds, 
                                  string mFileName, 
                                  string title, 
                                  out string errorString)
        {
            var h = CheckExcellProcesses();

            errorString = string.Empty;
            var excelApp = new Application
            {
                Visible = false,
                DisplayAlerts = false,
                DisplayClipboardWindow = false,
                DisplayFullScreen = false,
                ScreenUpdating = false,
                WindowState = XlWindowState.xlNormal
            };


            // Create an Excel workbook instance 
            var workbooks = excelApp.Workbooks;
            var excelWorkBook = workbooks.Add(Template: Type.Missing);

            var wst = excelWorkBook.Worksheets;
            Worksheet excelWorkSheet = null;
            const int startRow = 3;
            const int fRow = 1;
            const int fCol = 1;
            // Sheet name is limited to 32
            const int limit = 32;

            try
            {
                var tm = title.Split(';');
                var i1 = 0;


                var dataSource = DataAccess.GetResult("SELECT @@servername + '(' + DB_NAME() + ')';");
                var userName = DataAccess.GetResult("SELECT SYSTEM_USER");


                foreach (DataTable table in ds.Tables)
                {
                    // Add a new worksheet to workbook with the Data table name
                    excelWorkSheet = (Worksheet)wst.Add();

                    // Name is limited to 32 chars
                    excelWorkSheet.Name = table.TableName.PadRight(limit).Substring(0, limit).Trim();

                    var endIndex = table.Rows.Count + startRow ;

                    var localTitle = tm.Length > i1 ? tm[i1] : "Unknown";

                    //FormatFirstRow(
                    //    excelWorkSheet.Range[excelWorkSheet.Cells[fRow, fCol], excelWorkSheet.Cells[fRow, table.Columns.Count]],
                    //    localTitle);

                    Range newHeader = excelWorkSheet.Range[excelWorkSheet.Cells[fRow, fCol],
                        excelWorkSheet.Cells[fRow, table.Columns.Count]];


                    newHeader.Value2 = localTitle;
                    newHeader.Select(); // <-----necessary run time exception
                    newHeader.Merge(false);

                    newHeader.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    newHeader.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    newHeader.Font.Bold = true;
                    newHeader.Font.ColorIndex = 32;
                    newHeader.Font.Size = 14;

                    var border = newHeader.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    border.ColorIndex = 31;
                    border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;


                    //FormatSecondRow(
                    //    excelWorkSheet.Range[excelWorkSheet.Cells[fRow + 1, fCol], excelWorkSheet.Cells[fRow + 1, table.Columns.Count]],
                    //    mFileName, dataSource, userName);

                    var rngSecondRow = excelWorkSheet.Range[excelWorkSheet.Cells[fRow + 1, fCol],
                        excelWorkSheet.Cells[fRow + 1, table.Columns.Count]];
                    rngSecondRow.Value2 =
                        $"Made by : {userName}; on host : {Environment.MachineName}; date :" +
                        $" {DateTime.Now.ToString(Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern)} " +
                        $"{DateTime.Now.ToString(Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortTimePattern)}; " +
                        $"file name : {mFileName};data source : {dataSource}";


                    rngSecondRow.Select(); // <-----necessary run time exception
                    rngSecondRow.Merge(false);

                    rngSecondRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    rngSecondRow.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    rngSecondRow.Font.ColorIndex = 15;

                    var header = excelWorkSheet.Range[excelWorkSheet.Cells[startRow, 1],
                        excelWorkSheet.Cells[startRow, table.Columns.Count]];
                    header.Interior.ColorIndex = 34;

                    for (var i = 1; i < table.Columns.Count + 1; i++)
                    {
                        excelWorkSheet.Cells[startRow, i] = table.Columns[i - 1].ColumnName;

                        var newRng = excelWorkSheet.Range[excelWorkSheet.Cells[startRow + 1, i],
                            excelWorkSheet.Cells[endIndex, i]];
                        // format as datetime
                        if (table.Columns[i - 1].DataType == Type.GetType("System.DateTime"))
                            newRng.NumberFormat =
                                Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern;
                        // format as decimal ( integer )
                        else if (table.Columns[i - 1].DataType == Type.GetType("System.Decimal") ||
                                 table.Columns[i - 1].DataType == Type.GetType("System.Int32"))
                        {
                            if (table.Columns[i - 1].DataType == Type.GetType("System.Decimal"))

                                newRng.NumberFormat =
                                    "#" + Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator
                                         + "##0" + Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator + "00";
                            else
                                newRng.NumberFormat =
                                    "#" + Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator + "##0";

                            newRng.HorizontalAlignment = XlHAlign.xlHAlignRight;
                        }
                        // default format
                        else
                            newRng.NumberFormat = "@";
                    }




                    if (pasteRange == false)
                    {
                        // slow! do not use like this
                        for (var j = 0; j <= table.Rows.Count - 1; j++)
                        {
                            for (var k = 0; k <= table.Columns.Count - 1; k++)
                            {
                                excelWorkSheet.Cells[j + startRow + 1, k + 1] = table.Rows[j][k];
                            }
                        }
                    }
                    else
                    {
                        // using Value2 
                        excelWorkSheet.Range[excelWorkSheet.Cells[startRow + 1, 1], excelWorkSheet.Cells[endIndex, table.Columns.Count]].Value2 =
                            General.Convert(table);
                        if (Debugger.IsAttached)
                            Console.WriteLine("Setting value for table :"+  table.TableName);

                    }

                    excelWorkSheet.Activate();

                    // bug in Excel 
                    //excelWorkSheet.Application.ActiveWindow.SplitRow = startRow;
                    //excelWorkSheet.Application.ActiveWindow.SplitColumn = fCol;
                    //excelWorkSheet.Application.ActiveWindow.FreezePanes = true;


                    // apply auto filter
                    var firstRow = (Range)excelWorkSheet.Rows[startRow];

                    firstRow.AutoFilter(fRow,
                        Criteria1: Type.Missing,
                        Operator: XlAutoFilterOperator.xlAnd,
                        Criteria2: Type.Missing,
                        VisibleDropDown: true);



                    var newRng2 = excelWorkSheet.Range[excelWorkSheet.Cells[1, 1],
                        excelWorkSheet.Cells[endIndex, table.Columns.Count + 1]];
                    newRng2.Columns.AutoFit();
                    i1++;

                }

                ((Worksheet)excelApp.ActiveWorkbook.Sheets[excelApp.ActiveWorkbook.Sheets.Count]).Delete();


                excelWorkBook.SaveAs(Filename: mFileName,
                    FileFormat: XlFileFormat.xlOpenXMLWorkbook,
                    Password: Missing.Value,
                    WriteResPassword: Missing.Value,
                    ReadOnlyRecommended: false,
                    CreateBackup: false,
                    AccessMode: XlSaveAsAccessMode.xlNoChange,
                    ConflictResolution: XlSaveConflictResolution.xlUserResolution,
                    AddToMru: true,
                    TextCodepage: Missing.Value,
                    TextVisualLayout: Missing.Value,
                    Local: Missing.Value);
                excelWorkBook.Close();
                excelApp.Quit();

            }
            catch (Exception ex)
            {
                excelWorkBook.Close();
                excelApp.Quit();
                errorString = ex.Message;
            }
            finally
            {
                Marshal.ReleaseComObject(excelWorkSheet);
                Marshal.ReleaseComObject(wst);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(excelWorkBook);
                Marshal.ReleaseComObject(excelApp);
                // ReSharper disable once RedundantAssignment
                excelWorkBook = null;
                workbooks = null;
                // ReSharper disable once RedundantAssignment
                excelApp = null;
                GC.Collect();
                KillExcel(h);
            }

        }


        private static void FormatFirstRow(Range newRng, string title)
        {
            newRng.Value2 = title;
            newRng.Select(); // <-----necessary run time exception
            newRng.Merge(false);

            newRng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            newRng.VerticalAlignment = XlVAlign.xlVAlignCenter;
            newRng.Font.Bold = true;
            newRng.Font.ColorIndex = 32;
            newRng.Font.Size = 14;

            var border = newRng.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = 2d;
            border.ColorIndex = 31;
            border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;





        }

        private static void FormatSecondRow(Range rngSecondRow, string mFile, string dataSource, string userName)
        {
            rngSecondRow.Value2 =
                $"Made by : {userName}; on host : {Environment.MachineName}; date :" +
                $" {DateTime.Now.ToString(Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern)} " +
                $"{DateTime.Now.ToString(Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortTimePattern)}; " +
                $"file name : {mFile};data source : {dataSource}";
            rngSecondRow.Select(); // <-----necessary run time exception
            rngSecondRow.Merge(false);

            rngSecondRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rngSecondRow.VerticalAlignment = XlVAlign.xlVAlignCenter;
            rngSecondRow.Font.ColorIndex = 15;
        }
        private static Hashtable CheckExcellProcesses()
        {
            var myHashtable = new Hashtable();
            var allProcesses = Process.GetProcessesByName("excel");
            
            var iCount = 0;

            foreach (var excelProcess in allProcesses)
            {
                myHashtable.Add(excelProcess.Id, iCount);
                iCount = iCount + 1;
            }

            return myHashtable;
        }

        private static void KillExcel(Hashtable myHashtable)
        {
            var allProcesses = Process.GetProcessesByName("excel");

            // check to kill the right process
            foreach (var excelProcess in allProcesses)
            {
                if (myHashtable.ContainsKey(excelProcess.Id) == false)
                    excelProcess.Kill();
            }

            // ReSharper disable once RedundantAssignment
            allProcesses = null;
        }

    }
}