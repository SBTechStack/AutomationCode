using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Outlook;


namespace OutlookOperations
{
    public class MSOutlookOperations
    {
        private MSOutlook mSOutlook = new MSOutlook();
        private MSOutlooks mSOutlooks = new MSOutlooks();
        private static MSOutlookOperations _instance;
        private static readonly object _lock = new object();
        private bool sortorder = false;
        private Dictionary<string, string> GetValuePairsofMailIems { get; set; } = new Dictionary<string, string>();
        private List<string> entryIDs = new List<string>();

        public static MSOutlookOperations Instance
        {
            get
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new MSOutlookOperations();
                    }
                    return _instance;
                }
            }
        }

        public string GetAttachments
        {
            get; set;
        }
        public string SortOptions { get; set; }
        public bool SortOrder
        {
            get { return sortorder; }
            set
            {
                sortorder = value;
            }
        }
        public string MailFolderPath{get; set;}
        public string MailFilter { get; set; }
        public string MailSort { get; set; }
       

        public void OpenOutlook()
        {
            Application outlookApp = new Application();
            outlookApp.CreateItem(OlItemType.olMailItem);
            Marshal.ReleaseComObject(outlookApp);
            outlookApp = null;
        }

        public void CloseOutlook()
        {
            Application outlookApp = new Application();
            outlookApp.Quit();
            Marshal.ReleaseComObject(outlookApp);
            outlookApp = null;
        }

        public void Delete(string MailItemName)
        {
            mSOutlooks[MailItemName].Delete();
            mSOutlooks.Remove(mSOutlooks[MailItemName]);
        }

        public void SendAndReceive(int WaitTimeInSeconds)
        {
            MSOutlook.SendAndReceive(WaitTimeInSeconds);
        }

        public void MarkAsReadUnread(string MailItemName, bool IsUnread = false)
        {
            mSOutlooks[MailItemName].MarkAsReadUnread(IsUnread);
        }

        public void MoveToFolder(string MailItemName, string FolderPath)
        {
            mSOutlooks[MailItemName].MoveToFolder(FolderPath);
        }

        public void SaveAttachmentToFile(string MailItemName, string DownloadedPath)
        {
            if (mSOutlooks[MailItemName].AttachmentCount == 0)
                throw new System.Exception("No attachment found in the mail item");
            foreach (Microsoft.Office.Interop.Outlook.Attachment attachment in mSOutlook.mailItem.Attachments)
            {
                string Outputfilepaths = System.IO.Path.Combine(DownloadedPath, "_" + DateTime.Now.ToString("ddMMyyyyhhmmss")
                              + "_" + RemoveSpace(attachment.FileName));
                attachment.SaveAsFile(Outputfilepaths);
            }
        }

        public void SaveAttachmentToFile(string MailItemName, string DownloadedPath, string AttachmentName)
        {
            if (mSOutlooks[MailItemName].AttachmentCount == 0)
                throw new System.Exception("No attachment found in the mail item");
            foreach (Microsoft.Office.Interop.Outlook.Attachment attachment in mSOutlook.mailItem.Attachments)
            {
                if (attachment.FileName.Equals(AttachmentName))
                {
                    string Outputfilepaths = System.IO.Path.Combine(DownloadedPath, "_" + DateTime.Now.ToString("ddMMyyyyhhmmss")
                                  + "_" + RemoveSpace(attachment.FileName));
                    attachment.SaveAsFile(Outputfilepaths);
                }
            }
        }

        public void SendMail()
        {
            string[] atts = new string[0];


            MSOutlook newItem = new MSOutlook("shivanand_belagali@outlook.com");
            newItem.Subject = "mSOutlook.Subject";
            newItem.Body = "mSOutlook.Body";
            newItem.BodyFormat = BodyFormat.HTML;

            if (!string.IsNullOrEmpty(GetAttachments))
            {
                if (GetAttachments.Contains("|"))
                {
                    atts = GetAttachments.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                }
                else
                {
                    atts[0] = GetAttachments;
                }
            }

            if (atts.Count() > 0)
            {
                foreach (string att in atts)
                {
                    Attachment attObj = newItem.mailItem.Attachments.Add(att, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    string[] imageExtensions = { ".PNG", ".JPG", ".JPEG", ".BMP", ".GIF" };
                    if (Array.IndexOf(imageExtensions, System.IO.Path.GetExtension(att).ToUpperInvariant()) != -1)
                    {
                        attObj.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F", "image/jpeg");
                        attObj.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", System.IO.Path.GetFileName(att));
                    }
                }
            }

            newItem.Send();
            Marshal.ReleaseComObject(newItem.mailItem);
            newItem.mailItem = null;

        }
    
        public void ProcessMails(string MailBoxFolderPath)
        {
            MAPIFolder folder = MSOutlook.FindFolder(MailBoxFolderPath);
            Items mailItems = folder.Items;

            int count = mailItems.Count;

            if (string.IsNullOrEmpty(MailFilter))
            {
                MailFilter = "[UnRead] = true";
            }

            if (!string.IsNullOrEmpty(MailFilter))
            {
                mailItems = mailItems.Restrict(MailFilter);
            }

            // Sort
            if (!string.IsNullOrEmpty(SortOptions))
            {
                mailItems.Sort(SortOptions, SortOrder);
            }

            foreach (object obj in mailItems)
            {
                if (obj is MailItem)
                {
                    GetValuePairsofMailIems.Add(((MailItem)obj).EntryID, folder.StoreID);
                    entryIDs.Add(((MailItem)obj).EntryID);
                }
            }

            if (GetValuePairsofMailIems.Count == 0)
            {
                throw new System.Exception("No mail items found in the folder");
            }

            var keyValuePairs = GetValuePairsofMailIems.FirstOrDefault();

            if (keyValuePairs.Key == null)
            {
                throw new System.Exception("No mail items found in the folder");
            }

            string lDownloadedoutputPath = Path.Combine(Environment.CurrentDirectory, "DownloadAttachment");
            if (keyValuePairs.Key != null)
            {
                mSOutlook.mailItem = MSOutlook.GetMailItem(keyValuePairs.Key, keyValuePairs.Value);
                foreach (Microsoft.Office.Interop.Outlook.Attachment attachment in mSOutlook.mailItem.Attachments)
                {
                    string Outputfilepaths = System.IO.Path.Combine(lDownloadedoutputPath, "_" + DateTime.Now.ToString("ddMMyyyyhhmmss")
                        + "_" + RemoveSpace(attachment.FileName));
                    attachment.SaveAsFile(Outputfilepaths);
                }
            }



        }

        private static string RemoveSpace(string lines)
        {
            return Regex.Replace(lines, @"\s", "");
        }


    }
}
