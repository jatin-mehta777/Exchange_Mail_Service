using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Configuration;

namespace Exchange_Mail_Service
{
    public static class ReadMails
    {

        public static void GetMails()
        {
            try
            {
                //TimeSpan ts = new TimeSpan(0, -1, 0, 0);
                TimeSpan ts = new TimeSpan(0, -1, 0, 0);
                DateTime date = DateTime.Now.Add(ts);
                //ExchangeService exchange = new ExchangeService(ExchangeVersion.Exchange2010);
                ExchangeService exchange = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                //exchange.Credentials = new WebCredentials("Domain\\Username", @"Password");
                exchange.Credentials = new WebCredentials("MS\\abc", @"password");
                //Uri - https://mail.comanyname.com/EWS/Exchange.asmx
                exchange.Url = new Uri("https://mail.uhc.com/EWS/Exchange.asmx");
                exchange.TraceEnabled = true;
                exchange.TraceFlags = TraceFlags.All;
                // exchange.AutodiscoverUrl("myemail@company.com", RedirectionCallback); 
                SearchFilter.IsGreaterThanOrEqualTo filter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date);
                Service1.Log("calling service");
                FindItemsResults<Item> findResults = exchange.FindItems(WellKnownFolderName.Inbox, filter, new ItemView(50));

                bool RedirectionCallback(string url)
                {
                    return url.ToLower().StartsWith("https://");
                }
                foreach (Item item in findResults.Items)
                {
                    GetAttachmentsFromEmail(exchange, item.Id, item.Subject);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
                Service1.Log(ex.Message);
            }
            void GetAttachmentsFromEmail(ExchangeService service, ItemId itemId, String subject)
            {
                // Bind to an existing message item and retrieve the attachments collection.
                // This method results in an GetItem call to EWS.
                EmailMessage message = EmailMessage.Bind(service, itemId, new PropertySet(ItemSchema.Attachments));
                // Iterate through the attachments collection and load each attachment.

                foreach (Attachment attachment in message.Attachments)
                {
                    if (attachment is FileAttachment)
                    {

                        FileAttachment fileAttachment = attachment as FileAttachment;
                        // fileAttachment.Load("Put mails attachments" + fileAttachment.Name);
                        fileAttachment.Load("C:\\MailAttachments\\" + fileAttachment.Name);
                        //System.IO.FileInfo fi = new System.IO.FileInfo("C:\\Mails\\" + fileAttachment.Name);
                        //if (fi.Exists)
                        //{
                        //    // Move file with a new name. Hence renamed.  
                        //    fi.MoveTo(@"C:\\Mails\\" + subject);
                        //    Console.WriteLine("File Renamed.");
                        //}
                        Console.WriteLine("File attachment name: " + fileAttachment.Name);
                    }
                    else // Attachment is an item attachment.
                    {
                        ItemAttachment itemAttachment = attachment as ItemAttachment;
                        // Load attachment into memory and write out the subject.
                        // This does not save the file like it does with a file attachment.
                        // This call results in a GetAttachment call to EWS.
                        itemAttachment.Load();
                        Console.WriteLine("Item attachment name: " + itemAttachment.Name);
                    }
                }
            }

        }

    }
}
