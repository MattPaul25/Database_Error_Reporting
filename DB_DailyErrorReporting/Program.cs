using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DB_DailyErrorReporting
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, my name is DER, Daily Error Reporting");
            Console.WriteLine("starting, I run queries and send the results as a spreadsheet");
            var interpret = new InterpretText();        
           
        }
    }
    
    class InterpretText
    {       

        public InterpretText()
        {
            string fileName = ConfigurationManager.ConnectionStrings["Emails"].ConnectionString;
            readFile(fileName);
        }
        private void readFile(string fileName)
        {
            using (StreamReader sr = new StreamReader(fileName))
            {
                string currentLine = sr.ReadLine();
                while (currentLine != null)
                {
                    int firstPipe = TextUtils.Search(currentLine, "|") + 1;
                    string sqlFiles = currentLine.Substring(firstPipe);
                    string user = currentLine.Substring(0, firstPipe - 1);
                    var gen = new GenerateEmail(sqlFiles, user);

                    currentLine = sr.ReadLine();
                }
            }
        }
    }
}
