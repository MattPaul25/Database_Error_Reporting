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
        private string user;
        private string excelFolder;

        public GenerateEmail(string sqlString, string user)
        {
            // TODO: Complete member initialization
            this.sqlString = sqlString;
            this.user = user;
            this.excelFolder = ConfigurationManager.ConnectionStrings["wbLoc"].ConnectionString;
            CreateSpreadsheets();
            CreateEmail();
            DeleteFiles();
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
        private void CreateEmail()
        {
            var emlSend = new SendEmail(user, excelFolder);
        }
        private void DeleteFiles()
        {
            DirectoryInfo dir = new DirectoryInfo(excelFolder);
            foreach (FileInfo fi in dir.GetFiles())
            {
                fi.Delete();
            }
        }
        

    }
}
