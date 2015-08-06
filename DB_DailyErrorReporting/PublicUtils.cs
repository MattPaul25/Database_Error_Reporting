using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using ClosedXML;
using ClosedXML.Excel;


namespace DB_DailyErrorReporting
{
    public static class TextUtils
    {
        public static void Comment(string Comment = "")
        {
            //writes comment to console and logs it on log
            Console.WriteLine(Comment);
            using (StreamWriter sw = new StreamWriter("History.txt", true))
            {
                sw.WriteLine(Comment);
                sw.Close();
            }
        }
        public static string RightOf(string yourString, string yourMarker)
        {
            //method or function that pulls everything right of a unique Marker
            string newString = "";
            int stringLen = yourString.Length;
            int markLen = yourMarker.Length;
            if (stringLen > markLen)
            {
                int cnt = 0;

                for (int i = (stringLen - markLen); i > 0; i--)
                {
                    cnt = cnt + 1;
                    string temp = yourString.Substring(i, markLen);
                    if (temp == yourMarker)
                    {
                        newString = yourString.Substring(i + markLen, cnt - 1);
                        break;
                    }
                }
            }
            return newString;
        }
        public static int Search(string yourString, string yourMarker, int yourInst = 1, bool caseSensitive = true)
        {
            //returns the placement of a string in another string
            int num = 0;
            int currentInst = 1;
            //if optional argument, case sensitive is false convert string and marker to lowercase
            if (!caseSensitive) { yourString = yourString.ToLower(); yourMarker = yourMarker.ToLower(); }
            bool found = false;
            try
            {
                while (num < yourString.Length)
                {
                    string testString = yourString.Substring(num, yourMarker.Length);
                    if (testString == yourMarker)
                    {
                        if (currentInst == yourInst)
                        {
                            found = true;
                            break;
                        }
                        currentInst++;
                    }
                    num++;
                }
            }
            catch
            {
                num = 0;
            }
            num = found ? num : 0;
            return num;
        }
    }

    public static class CollectionUtils
    {
        public static void ConvertDataTableToExcel(System.Data.DataTable dt, string ExportLocation)
        {
            try
            {
                XLWorkbook wb = new XLWorkbook();
                dt.TableName = "Results";
                wb.Worksheets.Add(dt);
                wb.SaveAs(ExportLocation);
            }
            catch (Exception x)
            {
                Console.WriteLine(x.Message);
            }
        }
    }
}
