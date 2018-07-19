using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;


namespace SimpleTalkExcellAddin.Utils
{
    internal static class General
    {
        public static void FormatAsTable(
                                          Excel.Range sourceRange,
                                          string tableName,
                                          string tableStyleName,
                                          bool isSelected
            )
        {
            sourceRange.Worksheet.ListObjects.Add(SourceType: Excel.XlListObjectSourceType.xlSrcRange,
                    Source: sourceRange, LinkSource: Type.Missing, XlListObjectHasHeaders: Excel.XlYesNoGuess.xlYes,
                    Destination: Type.Missing).Name = tableName;

            if (isSelected)
                sourceRange.Select();

            sourceRange.Worksheet.ListObjects[tableName].TableStyle = tableStyleName;
        }

        internal static void AddChart(
            Excel.Worksheet pivotWorkSheet,
            string myTitle,
            Excel.Range pivotData,
            XlChartType type)
        {
            var chartObjects = (Excel.ChartObjects)pivotWorkSheet.ChartObjects();
            var pivotChart = chartObjects.Add(Left: 60, Top: 250, Width: 325, Height: 275);

            pivotChart.Chart.ChartWizard(Source: pivotData,
                Gallery: type,
                Title: myTitle,
                HasLegend: true,
                CategoryLabels: 6,
                SeriesLabels: 0);

            pivotChart.Chart.Location(Where: Excel.XlChartLocation.xlLocationAsNewSheet, Name: Type.Missing);
        }

        internal static void AddSlicers(
            Excel.Workbook eWorkbook,
            Excel.PivotTable pivotTable,
            Excel.Worksheet pivotWorkSheet,
            IEnumerable<string> slicerNames)
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
            IEnumerable<string> fieldRowNames,
            IEnumerable<string> fieldColumnNames,
            IEnumerable<string> fieldValueNames,
            IEnumerable<string> fieldReportNames
        )
        {
            var sheets = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets;
            var pivotTableName = myTitle;
            var pivotWorkSheet = (Excel.Worksheet)sheets.Add();
            pivotWorkSheet.Name = myTitle + (counter);


            //Add a PivotTable to the worksheet.
            pivotWorkSheet.PivotTableWizard(
                SourceType: Excel.XlPivotTableSourceType.xlDatabase,
                SourceData: pivotData,
                TableDestination: pivotWorkSheet.Cells[1, 1], // first cell of the pivot table
                TableName: pivotTableName
            );

            //Set variables used to manipulate the PivotTable.
            var pivotTable =
                (Excel.PivotTable)pivotWorkSheet.PivotTables(pivotTableName);


            //Format the PivotTable.
            pivotTable.TableStyle2 = tableStyle;
            pivotTable.InGridDropZones = false;
            pivotTable.GrandTotalName = "Total";


            foreach (var name in fieldRowNames)
            {
                var distValues = dataSource.DefaultView.ToTable(true, name).Rows.Count;
                if (distValues > 1000) continue;

                var rowField = (Excel.PivotField)pivotTable.PivotFields(name);
                rowField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

            }


            foreach (var name in fieldColumnNames)
            {
                // find out number of distinct values 
                var distValues = dataSource.DefaultView.ToTable(true, name).Rows.Count;
                if (distValues > 255) continue;
                var colField = (Excel.PivotField)pivotTable.PivotFields(name);
                colField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            }


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
