﻿using System;
using System.Configuration;
using System.Runtime.InteropServices.ComTypes;
using ExportToExcel.Data;

namespace ExportToExcel
{
    static class Program
    {
        // ReSharper disable once UnusedParameter.Local
        // ReSharper disable once ArrangeTypeMemberModifiers
        static void Main(string[] args)
        {
            var t1 = DateTime.Now;
            var query = ConfigurationManager.AppSettings["TestQuery"];
            const string name = "Sales_SalesOrderHeader;Sales_SalesOrderDetail";
            Console.WriteLine("Getting Data Source");
            var dataSet = DataAccess.GetDataSet(query,false,null, name , out var error);
            if (error != string.Empty)
            {
                Console.WriteLine(error);
                Console.ReadLine();
                return;
            }
            const string fileName = @"C:\TMP\FirstTest.xlsx";
            Console.WriteLine("Making Excel file");
            Excel.ExportToExcel.Export(true,dataSet, fileName, name, out var error2);
            if (error2 != string.Empty)
            {
                Console.WriteLine(error2);
                Console.ReadLine();
                return;
            }

            var t2 = DateTime.Now;
            TimeSpan ts = t2 - t1;

            Console.WriteLine($"Success in {ts}! Check file : {fileName}");
            Console.WriteLine("Press any key to exit");
            Console.ReadLine();

        }
    }
}