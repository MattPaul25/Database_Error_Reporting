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
        private string AttachmentDestination;
        
        public SendEmail(EmailObject eml, string DownloadDestination)
        {
            AttachmentDestination = DownloadDestination;
            today = DateTime.Today;            
            sendEmail(eml);
        }
        private void sendEmail(EmailObject eml)
        {
            string myDate = today.ToString("dd MMMM yyyy");
            //new outlook instance
            
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem);
            {
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
                    mail.Subject = myDate + " " + eml.EmailSubject;
                    mail.To = eml.Emails;
                    mail.Body = eml.EmailBody;
                    Console.WriteLine("sending...");
                    mail.Send();
                }
            }
            
        }
    }
}
