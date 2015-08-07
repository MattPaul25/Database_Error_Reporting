using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Configuration;
using System.IO;

namespace DB_DailyErrorReporting
{
    class GenerateEmail
    {
        private string sqlString;
        private string users;
        private string excelFolder;

        public GenerateEmail(EmailObject emlObj)
        {
            // TODO: Complete member initialization
            this.sqlString = emlObj.Queries;
            this.users = emlObj.Emails;
            this.excelFolder = ConfigurationManager.ConnectionStrings["wbLoc"].ConnectionString;
            CreateSpreadsheets();
            var emlSend = new SendEmail(emlObj, excelFolder);
    
        }

       
        private void CreateSpreadsheets()
        {
            string[] sqlFiles = sqlString.Split('|');
            foreach (string sqlFile in sqlFiles)
            {
                string sqlFilePath = ConfigurationManager.ConnectionStrings["SqlFiles"].ConnectionString + sqlFile;
                string sqlText = new MakeSQLString(sqlFilePath).sql;
                DataTable run = new RunSql(sqlText).sqlData;
                if (run.Rows.Count > 0)
                {
                    run.TableName = sqlFile;
                    string excelFileName = sqlFile.Substring(0, TextUtils.Search(sqlFile, ".sql")) + ".xlsx";
                    CollectionUtils.ConvertDataTableToExcel(run, excelFolder + excelFileName);
                }
            }
        }
             

    }
}
