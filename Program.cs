using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAttachmentExtractor
{
    class Program
    {
         static string basePath = @"c:\temp\sample1\";
         static int totalfilesize = 0;
        static void Main(string[] args)
        {
            EnumerateAccounts();
        }
        
        static void EnumerateFolders(Outlook.Folder folder)
        {
            Outlook.Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {                
                foreach (Outlook.Folder childFolder in childFolders)
                {                   
                    if (childFolder.FolderPath.Contains("Inbox"))
                    {                        
                        Console.WriteLine(childFolder.FolderPath);
                        EnumerateFolders(childFolder);
                    }
                }
            }
            Console.WriteLine("Looking for items in " + folder.FolderPath);
            IterateMessages(folder);
        }        
        static void IterateMessages(Outlook.Folder folder)
        {            
            string[] extensionsArray = {".pdf", ".doc",
".xls", ".ppt", ".vsd", ".zip",
".rar", ".txt", ".csv", ".proj" };
                       
            var fi = folder.Items;
            if (fi != null)
            {
                try
                {
                    foreach (Object item in fi)
                    {
                        Outlook.MailItem mi = (Outlook.MailItem)item;   
                        var attachments = mi.Attachments;
                       
                       // var mi.Attachments.Count = 15;
                                                
                            if (!System.IO.Directory.Exists(basePath + folder.FolderPath))
                            {
                                Directory.CreateDirectory(basePath + folder.FolderPath);
                            }
                                                      
                            for (int i = 1; i <= mi.Attachments.Count; i++)
                            {
                                var fn = mi.Attachments[i].FileName.ToLower();
                                if (extensionsArray.Any(fn.Contains))
                                {                                   
                                    if (!Directory.Exists(basePath + folder.FolderPath +
                                        @"\" + mi.Sender.Address))
                                    {
                                        Directory.CreateDirectory(basePath +
                                            folder.FolderPath + @"\" + mi.Sender.Address);
                                    }
                                    totalfilesize = totalfilesize + mi.Attachments[i].Size;
                                    if (!File.Exists(basePath + folder.FolderPath + @"\" +
                                        mi.Sender.Address + @"\" + mi.Attachments[i].FileName))
                                    {
                                        Console.WriteLine("Saving " + mi.Attachments[i].FileName);
                                        mi.Attachments[i].SaveAsFile(basePath + folder.FolderPath +
                                            @"\" + mi.Sender.Address + @"\" +
                        mi.Attachments[i].FileName);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Already saved " + mi.Attachments[i].FileName);
                                    }
                                }
                            }
                        }
                    
                }
                catch (Exception e)
                {
                    // Console.WriteLine("An error occurred: '{0}'", e);
                }
            }
        }

        static string EnumerateAccountEmailAddress(Outlook.Account account)
        {
            try
            {
                if (string.IsNullOrEmpty(account.SmtpAddress) || string.IsNullOrEmpty(account.UserName))
                {
                    Outlook.AddressEntry oAE = account.CurrentUser.AddressEntry as Outlook.AddressEntry;
                    if (oAE.Type == "EX")
                    {
                        Outlook.ExchangeUser oEU = oAE.GetExchangeUser() as Outlook.ExchangeUser;
                        return oEU.PrimarySmtpAddress;
                    }
                    else
                    {
                        return oAE.Address;
                    }
                }
                else
                {
                    return account.SmtpAddress;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "";
            }
        }

        static void EnumerateAccounts()
        {
            Console.Clear();
            Console.WriteLine("Outlook Attachment Extractor v0.1");
            Console.WriteLine("---------------------------------");
            int id;
            Outlook.Application Application = new Outlook.Application();
            Outlook.Accounts accounts = Application.Session.Accounts;

            string response = "";
            while (true == true)
            {
                id = 1;
                foreach (Outlook.Account account in accounts)
                {
                    Console.WriteLine(id + ":" + EnumerateAccountEmailAddress(account));
                    id++;
                }
                Console.WriteLine("Q: Quit Application");

                response = Console.ReadLine().ToUpper();
                if (response == "Q")
                {
                    Console.WriteLine("Quitting");
                    return;
                }
                if (response != "")
                {
                    if (Int32.Parse(response.Trim()) >= 1 && Int32.Parse(response.Trim()) < id)
                    {
                        Console.WriteLine("Processing: " +
                            accounts[Int32.Parse(response.Trim())].DisplayName);
                        Console.WriteLine("Processing: " +
                            EnumerateAccountEmailAddress(accounts[Int32.Parse(response.Trim())]));

                        Outlook.Folder selectedFolder =
                            Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
                        selectedFolder = GetFolder(@"\\" +
                            accounts[Int32.Parse(response.Trim())].DisplayName);
                        EnumerateFolders(selectedFolder);
                        Console.WriteLine("Finished Processing " +
                            accounts[Int32.Parse(response.Trim())].DisplayName);
                        Console.WriteLine("");
                    }
                    else
                    {
                        Console.WriteLine("Invalid Account Selected");
                    }
                }
            }
        }              
        static Outlook.Folder GetFolder(string folderPath)
        {
            Console.WriteLine("Looking for: " + folderPath);
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
                String[] folders = folderPath.Split(backslash.ToCharArray());
                Outlook.Application Application = new Outlook.Application();
                folder = Application.Session.Folders[folders[0]] as Outlook.Folder;
                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        Outlook.Folders subFolders = folder.Folders;
                        folder = subFolders[folders[i]] as Outlook.Folder;
                        if (folder == null)
                        {
                            return null;
                        }
                    }
                }
                return folder;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
    }
}
// }
//}
