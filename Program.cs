using Econocom.Data.Services;
using Econocom.Data.Services.Providers;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Configuration;
using Econocom.Core.Models.BlobStorages;
using Microsoft.Azure;
using System.IO;
using Econocom.Core.Models;

namespace OutlookEmailExtractor
{
    class Program
    {

        static string basePath = @"C:\Temp\emails\";
        static string SenderEmailAddress = "sales.op5@samsung.com";

        public static DatabaseSp sp;


        static void Main(string[] args)
        {
            Outlook.Application Application = new Outlook.Application();
            Outlook.Accounts accounts = Application.Session.Accounts;
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
            sp = new DatabaseSp(new ConnectionProvider(), new Econocom.Data.Services.Tools.GenericUtils());
            Outlook.MailItem newEmail = null;

            string DatabaseConnectionString = ConfigurationManager.ConnectionStrings["DataSourceDB"].ToString();
            string AzureStorageConnectionString = CloudConfigurationManager.GetSetting("AzureStorageConnectionString").ToString();
            string AzureStorageContainer = CloudConfigurationManager.GetSetting("AzureStorageContainer").ToString();
            var IDBlobStorageType = Convert.ToInt32(CloudConfigurationManager.GetSetting("IDBlobStorageTypeGeneric"));
            var BlobStorageRequest = new Econocom.Core.Models.BlobStorageRequest();
            var BlobStorageResponse = new Econocom.Core.Models.BlobStorageResponse();
            var clsBlobStorage = new BlobStorageService();
            var clsBlobStorageRequest = new BlobStorageUploadRequest();
            var clsBlobStorageResponse = new BlobStorageUploadResponse();
            string ReturnMessage = string.Empty;
            string line;
            //const string PR_ATTACH_DATA_BIN = "http://schemas.microsoft.com/mapi/proptag/0x37010102";
            var Login = new Econocom.Core.Models.Login();


       
            // SYS USER
            Login.IDLogin = 1;
            Login.IDWebSite = 2;
            Login.IDZone = 104;
            Login.IDSection = 1;
            Login.IDLanguage = "ENG";

            //Iterating each of mail item

            var fi = folder.Items;

            if (fi != null)
            {
                foreach (System.Object mail in fi)
                {
                    newEmail = mail as Outlook.MailItem;

                    if (newEmail != null && newEmail.UnRead == true && newEmail.SenderEmailAddress == SenderEmailAddress)
                    {
                        if (newEmail.Attachments.Count > 0)
                        {
                            for (int i = 1; i <= newEmail
                               .Attachments.Count; i++)
                            {
                                string[] extensionsArray = { ".txt" };

                                if (extensionsArray.Any(newEmail.Attachments[i].FileName.Contains))
                                {

                                    //newEmail.Attachments[i].SaveAsFile
                                    // (pathToSaveFile +
                                    // newEmail.Attachments[i].FileName);

                                    //Byte [] attachmentData = newEmail.Attachments[i].PropertyAccessor.GetProperty(PR_ATTACH_DATA_BIN);

                                    string tempFilePath = Path.GetTempPath() + newEmail.Attachments[i].FileName;
                                    FileStream fs = new FileStream(tempFilePath, FileMode.Create);
                                    var x = fs.Length;

                                    //StreamReader file = new StreamReader(tempFilePath);

                                    //var y = file.




                                    //  Blob storage file save
                                    //BlobStorageRequest.FileName = newEmail.Attachments[i].FileName;


                                    //BlobStorageRequest.FileSize = (attachmentData.Length / 1024);
                                    //BlobStorageRequest.ContentType = "text/plain";
                                    //BlobStorageRequest.IDBlobItem = null;
                                    //BlobStorageRequest.IDBlobStorageType = Convert.ToInt32(IDBlobStorageType);

                                    BlobStorageResponse = sp.BlobStorageAdd(Login, BlobStorageRequest);

                                    clsBlobStorage = new BlobStorageService();
                                    clsBlobStorageRequest = new BlobStorageUploadRequest();
                                    clsBlobStorageResponse = new BlobStorageUploadResponse();
                                    

                                    //ms.Write(attachmentData, 0, attachmentData.Length);
                                    //ms.Position = 0;
                                    //clsBlobStorageRequest.FileName = BlobStorageResponse.BlobStorage;
                                    //clsBlobStorageRequest.MemoryStream = ms;

                                    clsBlobStorageRequest.IDBlobStorageType = IDBlobStorageType.ToString();
                                    clsBlobStorageRequest.DatabaseConnectionString = ConfigurationManager.ConnectionStrings.ToString();
                                    clsBlobStorageRequest.StorageConnectionString = CloudConfigurationManager.GetSetting("AzureStorageConnectionString");
                                    clsBlobStorageRequest.StorageContainer = CloudConfigurationManager.GetSetting("AzureStorageContainer");

                                    clsBlobStorageResponse = clsBlobStorage.BlobUpload(clsBlobStorageRequest);
                                    if ((clsBlobStorageResponse.ReturnCode != "1"))
                                    {
                                        Console.WriteLine(clsBlobStorageResponse.ReturnMessage);
                                    }
                                    else
                                    {
                                        Console.WriteLine("OK");
                                    }

                                    // Read the file and save the information line by line.  
                                    //StreamReader file = new StreamReader(pathToSaveFile + newEmail.Attachments[i].FileName);

                                    //while ((line = file.ReadLine()) != null)
                                    //{
                                    //    if (line.Trim().Contains("PO No"))
                                    //    {
                                    //        int length = line.Length;
                                    //        var order = line.Substring(8, 10).Trim();
                                    //        var orderDate = line.Substring(26, 10).Trim().Replace(".","-");

                                    //    }

                                    //    if (line.Trim().Contains("Sales Person"))
                                    //    {
                                    //        int length = line.Length; //testing the length
                                    //        var DistributorCode = line.Substring(15, 14).Trim();

                                    //    }

                                    //    if (line.Trim().StartsWith("1"))
                                    //    {
                                    //        var ArticleCode = line.Substring(2, 14).Trim();
                                    //        var Quantity = line.Substring(16, 2).Trim();
                                    //        var Price = line.Substring(18, 5).Trim();
                                    //        var Delivery = line.Substring(33, 10).Trim().Replace(".", "-"); ;

                                    //    }

                                    //    if (line.Trim().StartsWith("Amount"))
                                    //    {
                                    //        var Amount = line.Substring(9, 4).Trim();

                                    //    }



                                    //}


                                    // Do the logic to call samsung order add spImportSamsungOrderAdd

                                    // if successful order add do the logic of calling spImportSamsungOrderJobSelect

                                    // If successfull then move the email in to imported folder


                                }

                            }

                            Console.WriteLine("Attachment downloaded successfully Please enter a key to close ");
                            Console.Read();

                        }
                    }
                }

            }
        }



        //static void MoveMailItem(Outlook.MailItem moveMail)
        //{

        //}

    }
}
