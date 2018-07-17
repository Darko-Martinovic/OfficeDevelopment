using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;


namespace SimpleTalkExcellAddin.Utils
{
    internal static class General
    {
        public static void FormatAsTable(Excel.Range sourceRange, string tableName, string tableStyleName)
        {
            sourceRange.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
                    sourceRange, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name =
                tableName;
            sourceRange.Select();
            sourceRange.Worksheet.ListObjects[tableName].TableStyle = tableStyleName;
        }

        internal static void AddChart(
            Excel.Worksheet pivotWorkSheet,
            string myTitle,
            Excel.Range pivotData,
            XlChartType type)
        {
            Excel.ChartObjects chartObjects = (Excel.ChartObjects)pivotWorkSheet.ChartObjects();
            Excel.ChartObject pivotChart = chartObjects.Add(Left: 60, Top: 250, Width: 325, Height: 275);

            //pivotChart.Chart.HasTitle = true;
            //pivotChart.Chart.ChartTitle.Text = "Ukupno";
            pivotChart.Chart.ChartWizard(pivotData,
                type,
                Title: myTitle,
                HasLegend: true,
                CategoryLabels: 6,
                SeriesLabels: 0);

            pivotChart.Chart.Location(Excel.XlChartLocation.xlLocationAsNewSheet, Type.Missing);
        }

        internal static void AddSlicers(
            Excel.Workbook eWorkbook,
            Excel.PivotTable pivotTable,
            Excel.Worksheet pivotWorkSheet,
            List<string> slicerNames)
        {
            const int top = 350;
            const int width = 200;
            const int height = 200;
            var left = 440;

            foreach (var name in slicerNames)
            {
                var slicerCurrent = eWorkbook.SlicerCaches.Add(pivotTable, name);

                slicerCurrent.Slicers.Add(pivotWorkSheet,
                    Top: top, Left: left, Width: width, Height: height);

                left += width + 10;
            }
        }

        internal static Excel.Worksheet AddPivot(
            string tableStyle,
            Excel.Range pivotData,
            DataTable dataSource,
            int counter,
            string myTitle,
            List<string> fieldRowNames,
            List<string> fieldColumnNames,
            List<string> fieldValueNames,
            List<string> fieldReportNames
        )
        {
            var sheets = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets;
            var pivotTableName = myTitle;

            //var sheets = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets;

            var pivotWorkSheet = (Excel.Worksheet)sheets.Add();
            pivotWorkSheet.Name = myTitle + (counter);

            // specify first cell for pivot table
            //Excel.Range oRange2 = pivotWorkSheet.Cells[1, 1];


            //Add a PivotTable to the worksheet.
            pivotWorkSheet.PivotTableWizard(
                Excel.XlPivotTableSourceType.xlDatabase,
                pivotData,
                pivotWorkSheet.Cells[1, 1], // first cell of the pivot table
                pivotTableName
            );

            ////Set variables used to manipulate the PivotTable.
            Excel.PivotTable pivotTable =
                (Excel.PivotTable)pivotWorkSheet.PivotTables(pivotTableName);


            //Format the PivotTable.
            //pivotTable.TableStyle2 = "PivotStyleLight16";
            pivotTable.TableStyle2 = tableStyle;
            pivotTable.InGridDropZones = false;
            //pivotTable.Summary = "Ukupno";
            //pivotTable.AlternativeText = "Tester";

            pivotTable.GrandTotalName = "Total";


            //pivotTable.AllowMultipleFilters = true; run-time exception
            //pivotTable.CompactLayoutColumnHeader = "Header";
            //pivotTable.Tag = "Tag";

            foreach (var name in fieldRowNames)
            {
                var distValues = dataSource.DefaultView.ToTable(true, name).Rows.Count;
                if (distValues > 1000) continue;

                var rowField = (Excel.PivotField)pivotTable.PivotFields(name);
                rowField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

            }

            //var i = 0;
            //Array.Resize(ref colField, fieldColumnNames.Length + 1);

            foreach (var name in fieldColumnNames)
            {
                // find out number of distinct values 
                var distValues = dataSource.DefaultView.ToTable(true, name).Rows.Count;
                if (distValues > 255) continue;


                var colField = (Excel.PivotField)pivotTable.PivotFields(name);

                colField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            }

            //colField2 = (Excel.PivotField)pivotTable.PivotFields("Kvartal");

            //colField2.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

            //colField3 = (Excel.PivotField)pivotTable.PivotFields("Mjesec");

            //colField3.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;


            //i++;

            //}



            foreach (var name in fieldValueNames)
            {
                var dataField = (Excel.PivotField)pivotTable.PivotFields(name);
                dataField.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                dataField.Function = Excel.XlConsolidationFunction.xlSum;

                dataField.NumberFormat = "#,##0.00"; //#.##0,00
            }



            foreach (var name in fieldReportNames)
            {
                var repField = (Excel.PivotField)pivotTable.PivotFields(name);
                repField.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
            }




            //colField.Caption = "Godina";
            return pivotWorkSheet;
        }

        internal static object[,] Convert(DataTable dt)
        {
            var rows = dt.Rows;
            var rowCount = rows.Count;
            var colCount = dt.Columns.Count;
            var result = new object[rowCount, colCount];

            for (var i = 0; i < rowCount; i++)
            {
                var row = rows[i];
                for (var j = 0; j < colCount; j++)
                {
                    result[i, j] = row[j];
                }
            }

            return result;
        }

        internal static string GetColumnName(int colNum)
        {
            var div = colNum;
            var name = string.Empty;

            while (div > 0)
            {
                var m = (div - 1) % 26;
                name = System.Convert.ToChar(65 + m) + name;
                div = (div - m) / 26;
            }

            return name;
        }
    }




}
