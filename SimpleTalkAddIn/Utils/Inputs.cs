using System.Collections.Generic;
using Microsoft.Office.Core;

namespace SimpleTalkExcellAddin.Utils
{
    public class Inputs
    {

        // ReSharper disable once CollectionNeverQueried.Global
        public List<string> AllFields { get; set; } = new List<string>();
        public List<string> ReportFielters { get; set; } = new List<string>();
        public List<string> Columns { get; set; } = new List<string>();
        public List<string> Rows { get; set; } = new List<string>();
        public List<string> Values { get; set; } = new List<string>();
        public string Query { get; set; }
        public string ConnectionString { get; set; }
        public XlChartType ChartType { get; set; }
        public string TableStyle { get; set; }
        public string PivotStyle { get; set; }
        //public  xls

        public static string GetConnectionString(string serverName, 
                                                 string dataBaseName, 
                                                 bool isWindowsAuth,
                                                 string userName,
                                                 string password)
        {
            var connectionString =
                $"Data Source={serverName};Integrated Security=SSPI;Initial Catalog={dataBaseName}";
            if (isWindowsAuth == false)
                connectionString =
                    $"Data Source={serverName};User Id={userName}; password= {password};Initial Catalog={dataBaseName}";
            return connectionString;
        }

    }
}
