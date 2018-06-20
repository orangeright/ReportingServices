using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;

namespace ReportingServices.Libraries
{
    public class DbAccess
    {
        private static SqlConnection GetSqlConnection()
        {
            return new SqlConnection(Parameters.Pleasanter.ConnectionString);
        }

        public static Stream GetTemplateFile(string guid)
        {
            string sql = "select Bin from [Binaries] where Guid = @guid";
            var conn = GetSqlConnection();
            var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.Add("@guid", SqlDbType.NVarChar).Value = guid;

            byte[] file = null;
            try
            {
                conn.Open();
                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    if (dr.HasRows)
                    {
                        dr.Read();
                        file = (byte[])dr[0];
                    }
                    else
                    {
                        // エラー
                    }
                }
            }
            catch(Exception ex)
            {

            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }

            return new MemoryStream(file);
        }


    }
}