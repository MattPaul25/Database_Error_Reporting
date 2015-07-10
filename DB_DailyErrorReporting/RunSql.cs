using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;

namespace DB_DailyErrorReporting
{
    class RunSql
    {
        public DataTable sqlData { get; set; }
        public RunSql(string sql)
        {
            sqlData = new DataTable();
            connect(sql);
        }
         private void connect(string sql)
         {
            
            string connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandTimeout = 2000;
                    sqlData.Load(cmd.ExecuteReader());

                }
            }
            catch (Exception x)
            {
                Console.WriteLine(x.Message);
            }

        }
    }
}
