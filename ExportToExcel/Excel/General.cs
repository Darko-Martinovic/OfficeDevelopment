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
    }
}