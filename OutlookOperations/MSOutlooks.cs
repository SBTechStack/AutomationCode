using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace OutlookOperations
{
    public abstract class NotifyProperyChangedBase : INotifyPropertyChanged
    {
        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;
        public bool IsCanvasValueChanged = false;

        protected bool CheckPropertyChanged<T>(string propertyName, ref T oldValue, ref T newValue)
        {
            if (oldValue == null && newValue == null)
            {
                IsCanvasValueChanged = false;
                return false;

            }

            if ((oldValue == null && newValue != null) || !oldValue.Equals((T)newValue))
            {
                oldValue = newValue;
                IsCanvasValueChanged = true;
                //FirePropertyChanged(propertyName);
                return true;
            }
            IsCanvasValueChanged = false;
            return false;
        }
        protected void FirePropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

    }
    public class MSOutlooks : List<MSOutlook>, INewProjectTransactions
    {
        //private Application outlookApp;
        //private NameSpace outlookNamespace;
        //private MAPIFolder inboxFolder;
        private string workbookFileName;
        private string workbookname = string.Empty;
        private string mProjectItemName = string.Empty;

        public string WorkbookName
        {
            get
            {
                workbookname = ProjectItemName;
                return ProjectItemName;
            }
            set
            {
                workbookname = value;
            }
        }
        public virtual string ProjectItemName
        {
            get { return mProjectItemName; }
            set
            {
                mProjectItemName = value;
            }
        }
        public string WorkbookFileName
        {
            get { return workbookFileName; }
            set { workbookFileName = value; }
        }
        public MSOutlook this[string index]
        {
            get
            {
                foreach (MSOutlook def in this)
                    if (def.MSOutlookName.Equals(index, StringComparison.CurrentCultureIgnoreCase)) return def;
                return null;
            }
        }
        public MSOutlook MSOutlook(string OutlookName)
        {
            return this[OutlookName];
        }
        public int IndexOf(string index)
        {
            if (this[index] == null) return -1;
            return this.IndexOf(this[index]);
        }
        public bool Contains(string index)
        {
            return (!(this[index] == null));
        }
        public bool Open(string FilePath)
        {
            throw new NotImplementedException();
        }
        public bool Save()
        {
            throw new NotImplementedException();
        }
        public bool SaveAs(string FilePath)
        {
            throw new NotImplementedException();
        }
        public bool Delete(string Objectname)
        {
            MSOutlook tProjectFlow = this.Where(f => f.ProjectItemID.Equals(Objectname)).FirstOrDefault();
            if (tProjectFlow != null)
            {
                return this.Remove(tProjectFlow);
            }
            else
                return false;
        }
        public bool Close()
        {
            throw new NotImplementedException();
        }
        public object AddNew(string Objectname)
        {
            throw new NotImplementedException();
        }
        public string New(string Objectname)
        {
            throw new NotImplementedException();
        }
        public object AddNew()
        {
            MSOutlook retValue = new MSOutlook();
            retValue.MSOutlookName = New();
            this.Add(retValue);
            return retValue;
        }
        public string New()
        {
            int i = 1;
            while (true)
            {
                if (this["MSOutlook" + i.ToString()] == null)
                    return "MSOutlook" + i.ToString();

                i++;
            }
        }
        public bool Rename(string OldObjectname, string NewObjectname)
        {
            MSOutlook tProjectFlow = this.Where(f => f.ProjectItemID.Equals(OldObjectname)).FirstOrDefault();
            if (tProjectFlow != null)
            {
                ProjectItemName = NewObjectname;
                return true;
            }
            else
                return false;
        }
    }
    public class MSOutlook : NotifyProperyChangedBase, IComparable
    {
        private Application outlookApp;
        private string mMSOutlookName;
        private bool IsNewMailItem = false;
        private string mProjectItemID = string.Empty;
        [XmlIgnore]
        public MailItem mailItem;
        public void OpenOutlook()
        {
            try
            {
                outlookApp = new Application();
                Marshal.ReleaseComObject(outlookApp);
                outlookApp = null;

            }
            catch (System.Exception ex)
            {
                throw new System.Exception("Unable to open Outlook", ex);
            }
        }
        public Application OutlookApp
        {
            get { return outlookApp; }
            set { outlookApp = value; }
        }
        public string ProjectItemID
        {
            get { return mProjectItemID; }
            set
            {
                mProjectItemID = value;
            }
        }

        public string MSOutlookName
        {
            get { return mMSOutlookName; }
            set { mMSOutlookName = value; }
        }

        [XmlIgnore]
        public string To
        {
            get
            {
                return mailItem.To;
            }
            set
            {
                if (this.IsNewMailItem)
                    mailItem.To = value;
            }
        }

        [XmlIgnore]
        public string FromName
        {
            get
            {
                return mailItem.Sender.Name;
            }
        }

        [XmlIgnore]
        public string FromEmailAddress
        {
            get
            {
                return mailItem.SenderEmailAddress;
            }
        }

        [XmlIgnore]
        public string Subject
        {
            get
            {
                return mailItem.Subject;
            }
            set
            {
                mailItem.Subject = value;
            }
        }

        [XmlIgnore]
        public string Body
        {
            get
            {
                return mailItem.Body;
            }
            set
            {
                mailItem.Body = value;
            }
        }

        [XmlIgnore]
        public string HTMLBody
        {
            get
            {
                return mailItem.HTMLBody;
            }
            set
            {
                mailItem.HTMLBody = value;
            }
        }

        [XmlIgnore]
        public BodyFormat BodyFormat
        {
            get
            {
                return (BodyFormat)Enum.Parse(typeof(BodyFormat), mailItem.BodyFormat.ToString().Replace("olFormat", ""));
            }
            set
            {
                switch (value)
                {
                    case BodyFormat.HTML:
                        mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
                        break;
                    case BodyFormat.RichText:
                        mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatRichText;
                        break;
                    case BodyFormat.Plain:
                        mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain;
                        break;
                    default:
                        mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatUnspecified;
                        break;
                }

                //mailItem.BodyFormat = (Microsoft.Office.Interop.Outlook.OlBodyFormat)
                //Enum.Parse(typeof(Microsoft.Office.Interop.Outlook.OlBodyFormat), value.ToString());


            }
        }

        [XmlIgnore]
        public DateTime ReceivedDateTime
        {
            get
            {
                return mailItem.ReceivedTime;
            }
        }

        [XmlIgnore]
        public DateTime SentTime
        {
            get
            {
                return mailItem.SentOn;
            }
        }

        public int AttachmentCount
        {
            get
            {
                return mailItem.Attachments.Count;
            }
        }

        public MailIAttachment Attachment(int Index)
        {
            return new MailIAttachment(mailItem.Attachments[Index]);
        }

        public void MoveToFolder(string FolderPath)
        {
            MAPIFolder mapiFolder = MSOutlook.FindFolder(FolderPath);
            if (mapiFolder != null)
            {
                mailItem.Move(mapiFolder);

                Marshal.ReleaseComObject(mapiFolder);
                mapiFolder = null;
            }
            else
            {
                throw new System.Exception("Outlook Folder [" + FolderPath + "] not found");
            }
        }

        public void Delete()
        {
            mailItem.Delete();
            Marshal.ReleaseComObject(mailItem);
            mailItem = null;
        }

        public void MarkAsReadUnread(bool markAsUnread)
        {
            mailItem.UnRead = markAsUnread;
        }
        public MSOutlook() { }
        public MSOutlook(string To)
        {
            Application app = null;
            try
            {
                app = new Application();

                this.mailItem = (MailItem)app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                this.mailItem.To = To;
            }
            finally
            {
                if (app != null)
                {
                    Marshal.ReleaseComObject(app);
                    app = null;
                }

                GC.Collect();
            }

            this.IsNewMailItem = true;
        }

        public void Send()
        {
            if (!this.IsNewMailItem)
                throw new System.Exception("Cannot send mail items not created in VW Studio");

            this.mailItem.Send();
        }

        public static MAPIFolder FindFolder(string path)
        {
            Application app = null;
            NameSpace MAPINamespace = null;
            MAPIFolder retValue = null;
            try
            {
                app = new Application();
                MAPINamespace = app.GetNamespace("MAPI");

                for (int i = 1; i <= MAPINamespace.Folders.Count; i++)
                {
                    string folderpath = MAPINamespace.Folders[i].FullFolderPath.Replace(@"\\", string.Empty).ToLower().Trim();
                    if (path.ToLower().Trim() == folderpath)
                    {
                        retValue = MAPINamespace.Folders[i];
                        break;
                    }
                    else if (path.ToLower().Trim().Equals(folderpath))
                    {
                        return FindFolder(path, MAPINamespace.Folders[i]);
                    }
                }

                return retValue;
            }
            finally
            {
                if (MAPINamespace != null)
                {
                    Marshal.ReleaseComObject(MAPINamespace);
                    MAPINamespace = null;
                }
                if (app != null)
                {
                    Marshal.ReleaseComObject(app);
                    app = null;
                }

                GC.Collect();
            }
        }

        private static MAPIFolder FindFolder(string path, MAPIFolder parent)
        {
            MAPIFolder retValue = null;

            for (int i = 1; i <= parent.Folders.Count; i++)
            {
                if (path == parent.Folders[i].FullFolderPath)
                {
                    retValue = parent.Folders[i];
                    break;
                }
                else if (path.StartsWith(parent.Folders[i].FullFolderPath))
                {
                    return FindFolder(path, parent.Folders[i]);
                }
            }

            return retValue;
        }

        public static MailItem GetMailItem(string entryID, string storeID)
        {
            Application app = null;
            NameSpace MAPINamespace = null;
            MAPIFolder retValue = null;
            try
            {
                app = new Application();
                MAPINamespace = app.GetNamespace("MAPI");
                return (MailItem)MAPINamespace.GetItemFromID(entryID, storeID);
            }
            finally
            {
                if (MAPINamespace != null)
                {
                    Marshal.ReleaseComObject(MAPINamespace);
                    MAPINamespace = null;
                }
                if (app != null)
                {
                    Marshal.ReleaseComObject(app);
                    app = null;
                }

                GC.Collect();
            }
        }

        public static void SendAndReceive(int WaitTimeInSeconds)
        {
            Application app = null;
            NameSpace MAPINamespace = null;

            try
            {
                app = new Application();
                MAPINamespace = app.GetNamespace("MAPI");
                MAPINamespace.SendAndReceive(false);
                System.Threading.Thread.Sleep(WaitTimeInSeconds * 1000);
            }
            finally
            {
                if (MAPINamespace != null)
                {
                    Marshal.ReleaseComObject(MAPINamespace);
                    MAPINamespace = null;
                }
                if (app != null)
                {
                    Marshal.ReleaseComObject(app);
                    app = null;
                }

                GC.Collect();
            }
        }

        public static string GetExchangeConnectionState()
        {
            Application app = null;
            NameSpace MAPINamespace = null;
            MAPIFolder retValue = null;
            try
            {
                app = new Application();
                MAPINamespace = app.GetNamespace("MAPI");
                return MAPINamespace.ExchangeConnectionMode.ToString();
            }
            finally
            {
                if (MAPINamespace != null)
                {
                    Marshal.ReleaseComObject(MAPINamespace);
                    MAPINamespace = null;
                }
                if (app != null)
                {
                    Marshal.ReleaseComObject(app);
                    app = null;
                }

                GC.Collect();
            }
        }

        public int CompareTo(object obj)
        {
            throw new NotImplementedException();
        }
    }
    public interface INewProjectTransactions
    {
        object AddNew(string Objectname);
        string New(string Objectname);
        object AddNew();
        string New();
        bool Rename(string OldObjectname, string NewObjectname);
    }
    public class MailIAttachment
    {
        public Attachment attItem;

        public string FileName
        {
            get
            {
                return attItem.FileName;
            }
        }

        public int Size
        {
            get
            {
                return attItem.Size;
            }
        }

        public void SaveAttachmentToFile(string saveFileName)
        {
            attItem.SaveAsFile(saveFileName);
        }

        public MailIAttachment(Attachment attItem)
        {
            this.attItem = attItem;
        }
    }
    public enum BodyFormat
    {
        HTML, RichText, Plain, Unspecified
    }
}
