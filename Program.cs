using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;


namespace GetAttachment
{
    class Program
    {
        static void Main(string[] args)
        {
            var service = new ExchangeService();
            service.Credentials = new NetworkCredential("krithika.shanmugakumar@philips.com", "Kittyapu*5");

            try
            {
                service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            }
            catch (AutodiscoverRemoteException ex)
            {
                Console.WriteLine(ex.Message);
            }

            FolderId inboxId = new FolderId(WellKnownFolderName.Inbox, "krithika.shanmugakumar@philips.com");
            var findResults = service.FindItems(inboxId, new ItemView(150));
            try
            {

                foreach (var message in findResults.Items)
                {

                    var msg = EmailMessage.Bind(service, message.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments));

                    foreach (Microsoft.Exchange.WebServices.Data.Attachment attachment in msg.Attachments)
                    {
                        if (attachment is FileAttachment)
                        {
                            FileAttachment fileAttachment = attachment as FileAttachment;

                            // Load the file attachment into memory and print out its file name.
                            fileAttachment.Load();
                            var filename = fileAttachment.Name;
                            bool b;
                            b = filename.Contains(".xlsx");

                            if (b == true)
                            {
                                bool a;
                                 bool k;
                                a = filename.Contains("Fields");
                                k = filename.Contains("MaterialID");
                                
                                    if (a == true || k== true)
                                    {
                                        var theStream = new FileStream("C:\\data\\kittu1\\" + fileAttachment.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                                        fileAttachment.Load(theStream);
                                        theStream.Close();
                                        theStream.Dispose();

                                    }
                                
                            }
                        }
                        else // Attachment is an item attachment.
                        {
                            // Load attachment into memory and write out the subject.
                            ItemAttachment itemAttachment = attachment as ItemAttachment;
                            itemAttachment.Load();
                        }

                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error Occured" + e.Message);
            }
            Console.WriteLine("Success");
            Console.ReadLine();
        }
    }
}

