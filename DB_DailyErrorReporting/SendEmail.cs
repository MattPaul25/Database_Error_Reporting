using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;

namespace DB_DailyErrorReporting
{
    class SendEmail
    {
        private DateTime today;
        public SendEmail(string DistributionList, string DownloadDestination)
        {
            today = DateTime.Today;
            string myDate = today.ToString("dd MMMM yyyy");
            sendEmail(DistributionList, DownloadDestination, myDate);
        }
        private void sendEmail(string DistributionList, string AttachmentDestination, string aDate)
        {           
            //new outlook instance
            Outlook.Application app = new Outlook.Application();            
            Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem);           
            string[] files = Directory.GetFiles(AttachmentDestination);
            int fileCount = 0;
            foreach (string file in files)
	        {
                Console.WriteLine("attatching file : " + file);
                mail.Attachments.Add(file);
                fileCount++;
	        }
            if (fileCount > 0)
            {
                mail.Importance = Outlook.OlImportance.olImportanceHigh;
                mail.Subject = aDate + "  Daily Report";
                mail.To = DistributionList;
                mail.Body = "Daily Report Email ";
                Console.WriteLine("sending...");
                mail.Send();
            }
        }
    }
}
