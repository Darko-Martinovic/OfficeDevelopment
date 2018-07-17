using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportFromExcel
{
    internal static class Program
    {
        private static void Main()
        {

            // Could be passed as parameter or read it from configuration file
            const string fileToRead = @"C:\TMP\FirstTest.xlsx";
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;

            try
            {
                xlApp = new Excel.Application();

                Console.WriteLine($"Trying to open file {fileToRead}");

                // Open Excel file and find out how many sheets are there, what are they names
                xlWorkBook = xlApp.Workbooks.Open(fileToRead, 0, true, 5, "", "", true,
                    Origin: Excel.XlPlatform.xlWindows, Delimiter: "\t", Editable: false, Notify: false, Converter: 0,
                    AddToMru: true, Local: 1, CorruptLoad: 0);

                var sheets = xlApp.ActiveWorkbook.Sheets;

                var dic = new List<string>();

                foreach (var mSheet in sheets)
                {
                    if (!(mSheet is Excel.Worksheet t1)) continue;

                    dic.Add($"[{t1.Name}$]");
                }

                // Could be read from configuration file
                using (var myConnection = new OleDbConnection($@"Provider=Microsoft.Ace.OLEDB.12.0;Data Source={fileToRead};Extended Properties='Excel 12.0 Xml;HDR = YES;'"))
                {
                    using (var dtSet = new DataSet())
                    {
                        foreach (var s in dic)
                        {
                            Console.WriteLine($" Processing {s} table");
                            var myCommand = new OleDbDataAdapter($@"select * from {s};", myConnection);
                            myCommand.TableMappings.Add("Table", s);
                            myCommand.Fill(dtSet);
                        }
                        foreach (DataTable t in dtSet.Tables)
                        {
                            Console.WriteLine($" Table {t.TableName} has {t.Rows.Count} records");
                        }
                    }
                }

                xlWorkBook.Close();
                xlApp.Quit();
                // ReSharper disable once RedundantAssignment
                dic = null;
                Console.WriteLine("Successfully imporeted!");
                Console.WriteLine("Press any key to exit");
                Console.ReadLine();

            }
            catch (Exception e)
            {
                xlWorkBook?.Close();
                xlApp?.Quit();
                Console.WriteLine($"Error importing from Excel : {e.Message}");
                Console.ReadLine();
            }
            finally
            {
                if (xlWorkBook != null)
                {
                    Marshal.ReleaseComObject(xlWorkBook);
                }
                if (xlApp != null)
                {
                    Marshal.ReleaseComObject(xlApp);
                }
                // ReSharper disable once RedundantAssignment
                xlWorkBook = null;
                // ReSharper disable once RedundantAssignment
                xlApp = null;
                GC.Collect();

            }



        }
    }
}
