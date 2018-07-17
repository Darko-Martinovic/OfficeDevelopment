using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace ExportToExcel.Data
{
    public static class DataAccess
    {
        public static DataSet GetDataSet(string query, bool isSp, SqlParameter[] listOfParams, string tableMapping, out string error)
        {
            var ds = new DataSet();
            error = string.Empty;
            var connectionString = ConfigurationManager.ConnectionStrings["ConnStr"].ConnectionString;
            try
            {
                using (var cnn = new SqlConnection(connectionString))
                {
                    using (var command = new SqlCommand(query, cnn))
                    {
                        cnn.Open();
                        if (isSp)
                        {
                            command.CommandType = CommandType.StoredProcedure;
                        }

                        if (listOfParams != null)
                        {
                            foreach (var p in listOfParams)
                            {
                                command.Parameters.Add(p);
                            }
                        }
                        var tm = tableMapping.Split(';');
                        var i = 0;
                        using (var sqlAdp = new SqlDataAdapter())
                        {
                            sqlAdp.SelectCommand = command;
                            foreach (var s in tm)
                            {
                                var addOn = i == 0 ? "" : i.ToString().Trim();
                                sqlAdp.TableMappings.Add("Table" + addOn, s);
                                i++;
                            }
                            sqlAdp.Fill(ds);
                        }
                        cnn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }
            return ds;
        }


        public static string GetResult(string query)
        {
            string ds;
            var connectionString = ConfigurationManager.ConnectionStrings["ConnStr"].ConnectionString;
            try
            {
                using (var cnn = new SqlConnection(connectionString))
                {
                    using (var command = new SqlCommand(query, cnn))
                    {
                        cnn.Open();
                        ds = (string) command.ExecuteScalar();
                        cnn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ds = ex.Message;
            }
            return ds;
        }

    }
}