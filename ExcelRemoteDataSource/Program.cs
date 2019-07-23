using System;
using System.Collections.Generic;
using System.Configuration;
using ExcelRemoteDataSource.ExcelHelper;


namespace ExcelRemoteDataSource
{
    public class Program
    {
        static void Main(string[] args)
        {
            var connectionString = ConfigurationManager.AppSettings["ConnStr"];
            var command = ConfigurationManager.AppSettings["Query"];
            var fieldRowsName = ConfigurationManager.AppSettings["Rows"].Split(',');
            var fieldValuesName = ConfigurationManager.AppSettings["Values"].Split(',');
            var fieldColumnsName = ConfigurationManager.AppSettings["Columns"].Split(',');
            var fieldReportFilter = ConfigurationManager.AppSettings["Reports"].Split(',');
            var fileName = ConfigurationManager.AppSettings["fileName"];
            var tableStyle = ConfigurationManager.AppSettings["PivotStyle"];
            var slicerStyle = ConfigurationManager.AppSettings["slicerStyle"];
            var chartType = ConfigurationManager.AppSettings["chartType"];
            var chartTitle = ConfigurationManager.AppSettings["chartTitle"];
            var sheetTitle = ConfigurationManager.AppSettings["sheetTitle"];
            var errorMessage = string.Empty;
            var eh = new ExcelHelper.ExcelHelper();
            try
            {
                eh.CreatePivotWithRemoteDataSource(
                    true,
                    true,
                    connectionString,
                    command,
                    fileName,
                    tableStyle,
                    chartType,
                    chartTitle,
                    sheetTitle,
                    fieldRowsName,
                    fieldValuesName,
                    fieldColumnsName,
                    fieldReportFilter,
                    fieldRowsName,
                    ref errorMessage,
                    slicerStyle);
                Console.WriteLine("Success!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            Console.ReadLine();
        }
    }
}
