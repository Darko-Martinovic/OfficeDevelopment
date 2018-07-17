using System.Data;

namespace ExportToExcel.Excel
{
    internal static class General
    {
       

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