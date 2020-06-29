using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookEmailExtractor
{
    class Program
    {

        static string basePath = @"C:\Temp\emails\";
        static string SenderEmailAddress = "sales.op5@samsung.com";
        static void Main(string[] args)
        {
            Outlook.Application Application = new Outlook.Application();
            Outlook.Accounts accounts = Application.Session.Accounts;
            //foreach (Outlook.Account account in accounts)
            //{
            //    Console.WriteLine(account.DisplayName );
            //}

            //Console.Read();

            Outlook.Folder root = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            EnumerateFolders(root);

        }

        static void EnumerateFolders(Outlook.Folder folder)
        {
            Outlook.Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    // We only want Posta in arrivo folders - ignore  others
                    if (childFolder.FolderPath.Contains("Posta in arrivo"))
                    {
                        //Console.WriteLine(childFolder.FolderPath);
                        // Call EnumerateFolders using childFolder, to see if there are any sub-folders within this one
                        //EnumerateFolders(childFolder);
                        IterateMessages(childFolder, basePath);
                    }
                }
            }
        }


        // This function can download the attachment of email which are sent by Samsung
        static void IterateMessages(Outlook.Folder folder, string pathToSaveFile)
        {
            Outlook.MailItem newEmail = null;
            var fi = folder.Items;

            if (fi != null)
            {

                foreach (Object mail in fi)
                {
                    newEmail = mail as Outlook.MailItem;

                    if (newEmail != null && newEmail.UnRead == true && newEmail.SenderEmailAddress == SenderEmailAddress)
                    {
                        if (newEmail.Attachments.Count > 0)
                        {
                            for (int i = 1; i <= newEmail
                               .Attachments.Count; i++)
                            {
                                newEmail.Attachments[i].SaveAsFile
                                    (pathToSaveFile +
                                    newEmail.Attachments[i].FileName);
                            }

                            Console.WriteLine("Attachment downloaded successfully Please enter a key to close ");
                            Console.Read();
                            
                        }
                    }
                }

            }
        }

    }
}
