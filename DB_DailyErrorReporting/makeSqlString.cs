using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Configuration;

namespace DB_DailyErrorReporting
{
    class MakeSQLString
    {
        private string sqlFile;
        public string sql { get; set; }

        public MakeSQLString(string sqlFile)
        {
            
            this.sqlFile = ConfigurationManager.ConnectionStrings["SqlFiles"].ConnectionString + sqlFile;
            sql = readSql(sqlFile);
        }

        private string readSql(string sqlFile)
        {
            string fullLine = "";
            using (StreamReader sr = new StreamReader(sqlFile))
            {
                string currentLine = sr.ReadLine();
                while (currentLine != null)
                {
                    fullLine += currentLine + "\n";
                    currentLine = sr.ReadLine();
                }
            }
            return fullLine;
        }
    }
}
