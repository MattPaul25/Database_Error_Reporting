using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Collections;
using System.Reflection;

namespace DB_DailyErrorReporting
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, my name is DER, Daily Error Reporting");
            Console.WriteLine("starting, I run queries and send the results as a spreadsheet");
            //var interpret = new InterpretText(); 
            string excelLocation = ConfigurationManager.ConnectionStrings["ExcelLocation"].ConnectionString; 
            var import = new ImportExcelTable(excelLocation, 1);
            if (import.isSuccessful)
            {
                var interpret = new InterpretDataTable(import.DataImport);
            }
        }
    }

    class ImportExcelTable
    {
        //takes downloaded file and imports it into memory
        private string ExcelFileLocation;
        public DataTable DataImport { get; protected set; }
        public bool isSuccessful { get; protected set; }

        public ImportExcelTable(string excelFileLocation, int SheetNum)
        {
            TextUtils.Comment("Importing Excel Data");
            DataImport = new DataTable();
            this.ExcelFileLocation = excelFileLocation;
            if (File.Exists(ExcelFileLocation))
            {
                convertExcelToDataTable(1);
                makeHeaders();
            }
            else
            {
                isSuccessful = false;
            }
        }

        private void convertExcelToDataTable(int worksheetNumber = 1)
        {
            if (!File.Exists(ExcelFileLocation)) throw new FileNotFoundException(ExcelFileLocation);

            // connection string
            var cnnStr = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO\"", ExcelFileLocation);

            using (var cnn = new OleDbConnection(cnnStr))
            {
                var dt = new DataTable();
                try
                {
                    cnn.Open();
                    var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (schemaTable.Rows.Count < worksheetNumber) throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet");
                    string worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
                    string sql = String.Format("select * from [{0}]", worksheet);
                    var da = new OleDbDataAdapter(sql, cnn);
                    da.Fill(dt);
                    DataImport = dt;
                    isSuccessful = true;
                }
                catch (Exception e)
                {
                    TextUtils.Comment(e.Message);
                    isSuccessful = false;
                }
                finally
                {
                    cnn.Close();
                }
            }
        }
        private void makeHeaders()
        {
            foreach (DataColumn column in DataImport.Columns)
            {
                string cName = DataImport.Rows[0][column.ColumnName].ToString();
                if (!DataImport.Columns.Contains(cName) && cName != "")
                {
                    column.ColumnName = cName;
                }
            }
            DataImport.Rows[0].Delete(); //Delete the row that has the headers
            DataImport.AcceptChanges();
        }
    }
    class InterpretDataTable
    {

        public InterpretDataTable(DataTable dt)
        {
            readTable(dt);
        }

        private void readTable(DataTable dt)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                    //convert object array to string array with linq expression
                    string[] emailData = ((IEnumerable)dt.Rows[i].ItemArray).Cast<object>()
                                                 .Select(x => x.ToString()).ToArray();
                    var emlObj = new EmailObject(emailData);
                    var gen = new GenerateEmail(emlObj);
            }
        }
    }

    class EmailObject
    {
        public string Emails { get; set; }
        public string EmailSubject { get; set; }
        public string EmailBody { get; set; }
        public string Queries { get; set; }
        
        public EmailObject(string[] emailData)
        {
            int i = 0;
            foreach (PropertyInfo prop in this.GetType().GetProperties())
            {
                prop.SetValue(this, emailData[i]);
                i++;
            }
        }
    }

}
