using System.Data;
using System.Data.SqlClient;

namespace SimpleTalkExcellAddin.Utils
{
    static class Helper
    {
        public static DataTable GetSource(Inputs i)
        {
            using (var dsData = new DataSet())
            {

                using (var conn = new SqlConnection(i.ConnectionString))
                {
                    var daData = new SqlDataAdapter(i.Query, conn);
                    daData.Fill(dsData);
                }
                return dsData.Tables[0];
            }
        }

    }
}
