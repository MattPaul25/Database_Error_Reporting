using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace DB_DailyErrorReporting
{
    class CleanUp
    {
        public string excelFolder { get; set; }
        public CleanUp()
        {
            this.excelFolder = ConfigurationManager.ConnectionStrings["wbLoc"].ConnectionString;
            DeleteFiles();
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
