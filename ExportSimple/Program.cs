using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportSimple
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            Excel.Workbooks wkb = xlApp.Workbooks;
            xlWorkBook = wkb.Add(misValue);
            Excel.Sheets wst = xlWorkBook.Worksheets;
            xlWorkSheet = (Excel.Worksheet)wst.Item[1];

            xlWorkSheet.Cells[1, 1] = "ID";
            xlWorkSheet.Cells[1, 2] = "Name";
            xlWorkSheet.Cells[2, 1] = "1";
            xlWorkSheet.Cells[2, 2] = "One";
            xlWorkSheet.Cells[3, 1] = "2";
            xlWorkSheet.Cells[3, 2] = "Two";



            xlWorkBook.SaveAs("d:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();


         



            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(wst);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(wkb);
            Marshal.ReleaseComObject(xlApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Console.WriteLine("Success");
            Console.ReadLine();

        }
    }
}
