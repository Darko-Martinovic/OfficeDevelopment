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
            Excel.Workbooks workbooks = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Sheets sheets = null;
            try
            {
                xlApp = new Excel.Application();

                Console.WriteLine($"Trying to open file {fileToRead}");

                // Open Excel file and find out how many sheets are there, what are they names
                workbooks = xlApp.Workbooks;

                xlWorkBook = workbooks.Open(fileToRead, 0, true, 5, "", "", true,
                    Origin: Excel.XlPlatform.xlWindows, Delimiter: "\t", Editable: false, Notify: false, Converter: 0,
                    AddToMru: true, Local: 1, CorruptLoad: 0);

                sheets = xlApp.ActiveWorkbook.Sheets;

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
                            // If an exception is thrown, probably you have to install a patch from the following link
                            // https://www.microsoft.com/en-us/download/confirmation.aspx?id=13255

                            myCommand.Fill(dtSet);
                        }
                        foreach (DataTable t in dtSet.Tables)
                        {
                            Console.WriteLine($" Table {t.TableName} has {t.Rows.Count} records");
                        }
                    }
                }



                // ReSharper disable once RedundantAssignment
                dic = null;
                Console.WriteLine("Successfully imported!");
                Console.WriteLine("After closing Console Windows start Task Manager and be sure that Excel instance is not there!");
                Console.WriteLine("Press any key to exit");
                Console.ReadLine();

            }
            catch (Exception e)
            {
                Console.WriteLine($"Error importing from Excel : {e.Message}");
                Console.ReadLine();
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (sheets != null)
                {
                    Marshal.FinalReleaseComObject(sheets);
                    sheets = null;
                }
                 

                xlWorkBook.Close();

                if (xlWorkBook != null)
                {
                    //Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.FinalReleaseComObject(xlWorkBook);
                    xlWorkBook = null;
                }

                xlApp.Quit();
                if (xlApp != null)
                {
                    //Marshal.ReleaseComObject(xlApp);
                    Marshal.FinalReleaseComObject(xlApp);
                    xlApp = null;

                }
                GC.Collect();
                GC.WaitForPendingFinalizers();


            }



        }
    }
}
