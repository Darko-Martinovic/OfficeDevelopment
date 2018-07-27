using System;
using System.Configuration;
using System.IO;
using ExportToExcel.Data;

namespace ExportToExcel
{
    static class Program
    {
        // ReSharper disable once UnusedParameter.Local
        // ReSharper disable once ArrangeTypeMemberModifiers
        static void Main(string[] args)
        {
            try
            {
                //Check if Excel is installed
                var officeType = Type.GetTypeFromProgID("Excel.Application");
                if (officeType == null)
                {
                    Console.WriteLine("Sorry, Excel must be installed!");
                    Console.WriteLine("Press any key to exit");
                    Console.ReadLine();
                    return;
                }
                var t1 = DateTime.Now;
                var query = ConfigurationManager.AppSettings["TestQuery"];
                const string name = "Sales_SalesOrderHeader;Sales_SalesOrderDetail";
                Console.WriteLine("Getting Data Source");
                var dataSet = DataAccess.GetDataSet(query, false, null, name, out var error);
                if (error != string.Empty)
                {
                    Console.WriteLine(error);
                    Console.ReadLine();
                    return;
                }
                const string fileName = @"C:\TMP\FirstTest.xlsx";
                Console.WriteLine("Making Excel file");
                //Create C:\TMP if not exists
                if (Directory.Exists(Path.GetDirectoryName(fileName)) == false)
                    Directory.CreateDirectory(Path.GetDirectoryName(fileName) ?? throw new InvalidOperationException());

                Excel.ExportToExcel.Export(pasteRange: true, ds: dataSet, mFileName: fileName, title: name,
                    errorString: out var error2);

                if (error2 != string.Empty)
                {
                    Console.WriteLine(error2);
                    Console.ReadLine();
                    return;
                }

                var t2 = DateTime.Now;
                var ts = t2 - t1;

                Console.WriteLine($"Success in {ts}! Check file : {fileName}");
                Console.WriteLine("After closing Console Windows start Task Manager and be sure that Excel instance is not there!");

                Console.WriteLine("Press any key to exit");
                Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadLine();
            }


        }
    }
}
