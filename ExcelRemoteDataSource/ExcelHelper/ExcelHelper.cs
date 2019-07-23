using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ExcelRemoteDataSource.ExcelHelper
{
	public class ExcelHelper

	{

		private Hashtable _myHashtable;


		public bool CreatePivotWithRemoteDataSource(
													bool showData, 
													bool spetialStyles,
													string connectionString, 
													string tsqlCommand, 
													string fileName, 
													string tableStyle, 
													string charType,
													string chartTitle, 
													string sheetTitle, 
													IEnumerable<string> pivotRowFields,
													IEnumerable<string> pivotValueFields, 
													IEnumerable<string> pivotColumnFields,
													IEnumerable<string> pivotReportFields,
													IEnumerable<string> slicerFieldName,
													ref string errorMessage,
													string slicerStyle, 
													string connectionName = "Pivot", 
													string connectionDescription = "Description",
													bool refreshOnOpen = false, 
													bool isIntegrated = false, 
													bool backQuery = true,
													bool createModelConnection = true, 
													bool importRelation = true)
		{
			Processes();
			var retValue = true;


			var excelApp = new Application()
			{
				Visible = false,
				DisplayAlerts = false,
				DisplayClipboardWindow = false,
				DisplayFullScreen = false,
				ScreenUpdating = false,
				WindowState = XlWindowState.xlNormal
			};

			var excelWorkBook = excelApp.Workbooks.Add(Template: Type.Missing);

			try
			{
				if (showData)
				{
					var isError = false;
					for (var i = 0; i <= 1; i++)
					{
						if (i == 1)
						{
							createModelConnection = false;
							importRelation = false;
						}
						try
						{
							var con = excelWorkBook.Connections.Add2(
								connectionName,
								connectionDescription,
								connectionString,
								tsqlCommand,
								XlCmdType.xlCmdSql,
								createModelConnection,
								importRelation);


							// con.OLEDBConnection.RefreshDate
							con.OLEDBConnection.RefreshOnFileOpen = refreshOnOpen;
							if (isIntegrated)
								con.OLEDBConnection.ServerCredentialsMethod = XlCredentialsMethod.xlCredentialsMethodIntegrated;
							else
							{
								con.OLEDBConnection.ServerCredentialsMethod = XlCredentialsMethod.xlCredentialsMethodStored;
								con.OLEDBConnection.SavePassword = true;
							}
							con.OLEDBConnection.BackgroundQuery = backQuery;
						}
						catch (Exception ex)
						{
							if (Debugger.IsAttached)
								Debugger.Break();
							isError = true;
						}
						if (isError == false)
							break;
					}
				}

				var pivotCache = excelApp.ActiveWorkbook.PivotCaches().Create(XlPivotTableSourceType.xlExternal,
																  Type.Missing,
																  XlPivotTableVersionList.xlPivotTableVersion15);

				pivotCache.Connection = connectionString;

				pivotCache.MaintainConnection = true;

				pivotCache.CommandText = tsqlCommand;

				pivotCache.CommandType = XlCmdType.xlCmdSql;

				Worksheet sheet = excelApp.ActiveSheet;


				sheet.Name = sheetTitle.Equals(string.Empty) ? "Not specified" : sheetTitle;

				PivotTables pivotTables = sheet.PivotTables();

				var cell = excelApp.ActiveCell;
				var slicerfieldNames = slicerFieldName.ToList();
				if (slicerFieldName != null && slicerfieldNames.Any())
					cell = sheet.Cells[5, 1];


				var pivotTable = pivotTables.Add(pivotCache,
									 cell,
									 "PivotTable1",
									 true,
									 XlPivotTableVersionList.xlPivotTableVersion15);

				AddPivotRemotely(
					excelWorkBook,
					sheet,
					pivotTable,
					pivotRowFields,
					pivotColumnFields,
					pivotValueFields,
					pivotReportFields,
					slicerfieldNames,
					slicerStyle);

				if (charType != string.Empty)
					AddChart2(
						sheet,
						chartTitle,
						pivotTable.TableRange1,
						(XlChartType)Enum.Parse(typeof(XlChartType), charType));



				sheet.Select();
				if (spetialStyles)
				{
					pivotTable.TableStyle2 = tableStyle;
					var rng = (Range)sheet.Cells[5, 1];
					rng.Select();
				}
				else
					pivotTable.TableStyle2 = tableStyle;

				try
				{
					excelWorkBook.Connections["Connection"].Description = connectionDescription;
				}
				catch (Exception ex)
				{
					if(Debugger.IsAttached)
						Debugger.Break();
				}


				excelWorkBook.SaveAs(fileName,
						 XlFileFormat.xlOpenXMLWorkbook,
						 Type.Missing,
						 Type.Missing,
						 false,
						 false,
						 XlSaveAsAccessMode.xlNoChange,
						 XlSaveConflictResolution.xlUserResolution,
						 true,
						 Type.Missing,
						 Type.Missing,
						 Local: Type.Missing);

				excelWorkBook.Close();

				excelApp.Quit();
			}

			catch (Exception ex)
			{
				retValue = false;
				errorMessage = ex.Message;
			}

			finally
			{
				Marshal.ReleaseComObject(excelWorkBook);

				Marshal.ReleaseComObject(excelApp);

				excelWorkBook = null;

				excelApp = null;
				GC.Collect();

				KillExcel();
			}
			return retValue;
		}

		private static void AddPivotRemotely(
			                           _Workbook eWorkBook, 
			                           Worksheet pivotWorkSheet, 
			                           PivotTable pivotTable,
			                           IEnumerable<string> fieldRowNames,
			                           IEnumerable<string> fieldColumnNames,
			                           IEnumerable<string> fieldValueNames,
			                           IEnumerable<string> fieldReportNames,
			                           IEnumerable<string> slicerfieldNames, 
			                           string slicerStyle)
		{
			PivotField rowField;
			PivotField dataField;
			PivotField colField;
			PivotField repField;


			pivotTable.InGridDropZones = false;
			pivotTable.GrandTotalName = "Total";
			pivotTable.DisplayNullString = false;
			pivotTable.ShowTableStyleRowHeaders = true;
			pivotTable.ShowTableStyleColumnHeaders = true;


			string masterFieldName = string.Empty;
			string masterColumnName = string.Empty;
			int rowCounter = 0;
			if (fieldRowNames != null)
			{
				foreach (var name in fieldRowNames)
				{
					rowCounter += 1;
					rowField = (PivotField)pivotTable.PivotFields(name);
					if (masterFieldName.Equals(string.Empty))
					{
						var pf = (PivotField) pivotTable.PivotFields(name);
						masterFieldName = pf.Name;
					}
						
					rowField.Orientation = XlPivotFieldOrientation.xlRowField;
				}
			}


			if (fieldColumnNames != null)
			{
				foreach (var name in fieldColumnNames)
				{
					colField = (PivotField)pivotTable.PivotFields(name);
					colField.Orientation = XlPivotFieldOrientation.xlColumnField;
					if (masterColumnName.Equals(string.Empty))
					{
						var pf = (PivotField) pivotTable.PivotFields(name);
						masterColumnName = pf.Name;
					}

						
				}
			}

			var nameOfSum = string.Empty;
			if (fieldValueNames != null)
			{
				foreach (var name in fieldValueNames)
				{
					dataField = (PivotField)pivotTable.PivotFields(name);
					dataField.Orientation = XlPivotFieldOrientation.xlDataField;
					dataField.Function = XlConsolidationFunction.xlSum;
					dataField.NumberFormat = "#,##0.00";
					dataField.Name = $"Total {name}";
					if (nameOfSum.Equals(string.Empty))
						nameOfSum = dataField.Name;
				}
			}

			if (!nameOfSum.Equals(string.Empty) && !masterFieldName.Equals(string.Empty))
			{
				// Sort DESC po prvoj numeričkoj veličini
				var pf = (PivotField) pivotTable.PivotFields(masterFieldName);
				pf.AutoSort((int)XlSortOrder.xlDescending, nameOfSum);

				if (rowCounter > 1)
				{
					var pf1 = (PivotField)pivotTable.PivotFields(masterFieldName);
					pf1.ShowDetail = false;
				}
					

				pivotTable.CompactLayoutRowHeader = masterFieldName;
			}

	

			if (fieldReportNames != null)
			{
				foreach (var name in fieldReportNames)
				{
					repField = (PivotField)pivotTable.PivotFields(name);
					repField.Orientation = XlPivotFieldOrientation.xlPageField;
				}
			}
			if (slicerfieldNames != null)
			{
				const int cTop = 300;
				const int cWidth = 100;
				const int cHeight = 100;

				var dLeft = 0;

				foreach (var name in slicerfieldNames)
				{
					var slicerCurrent = eWorkBook.SlicerCaches.Add(pivotTable, name);


					var slicer = slicerCurrent.Slicers.Add(
						pivotWorkSheet,
						Top: cTop,
						Left: dLeft,
						Width: cWidth * slicerCurrent.SlicerItems.Count,
						Height: cHeight);


					slicer.NumberOfColumns = slicerCurrent.SlicerItems.Count/2;
					slicer.Style = slicerStyle;


					dLeft += cWidth + 20;
				}
			}
		}


		private static void AddChart2(_Worksheet pivotWorkSheet, string myTitle, Range pivotData, XlChartType type)
		{
			var chartObjects = (ChartObjects)pivotWorkSheet.ChartObjects();

			var pivotChart = chartObjects.Add(Left: 60, Top: 250, Width: 325, Height: 275);


			pivotChart.Chart.SetSourceData(pivotData);
			pivotChart.Chart.ChartType = type;
			if (myTitle.Equals(string.Empty) == false)
				pivotChart.Name = myTitle;
			pivotChart.Chart.Location(XlChartLocation.xlLocationAsNewSheet);
		}

		private void KillExcel()
		{
			var allProcesses = Process.GetProcessesByName("excel");

			foreach (var excelProcess in allProcesses)
			{
				if (_myHashtable.ContainsKey(excelProcess.Id) == false)
					excelProcess.Kill();
			}

			allProcesses = null;
		}

		private void Processes()
		{
			var allProcesses = Process.GetProcessesByName("excel");
			_myHashtable = new Hashtable();
			var iCount = 0;

			foreach (var excelProcess in allProcesses)
			{
				_myHashtable.Add(excelProcess.Id, iCount);
				iCount += 1;
			}
		}


	}
}
