using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Xml.Serialization;
using System.Reflection;
using System.Drawing;
using System.Runtime.InteropServices;
 
using System.Data;
 
using System.ComponentModel;

namespace ExcelOperations
{
    public class ExcelWorkbooks : List<ExcelWorkbook>, INotifyPropertyChanged
    {

        public event PropertyChangedEventHandler PropertyChanged;
        public ExcelWorkbooks()
        {
        }
        public ExcelWorkbook this[string index]
        {
            get
            {
                foreach (ExcelWorkbook def in this)
                {
                    if(def.WorkbookName.Equals(index.ToString(), StringComparison.CurrentCultureIgnoreCase))
                    {
                        return def;
                    }   
                }                 
                return null;
            }
        }
        public ExcelWorkbook WkBook(string WorkbookName)
        {
            return this[WorkbookName] as ExcelWorkbook;

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
            ExcelWorkbook tProjectFlow = this.Where(f => f.ProjectItemID.Equals(Objectname)).FirstOrDefault();
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
        public object AddNew()
        {
            ExcelWorkbook retValue = new ExcelWorkbook();
            retValue.WorkbookName = New();
            this.Add(retValue);
            return retValue;
        }
        public string New()
        {
            int i = 1;
            while (true)
            {
                if (this["Workbook" + i.ToString()] == null)
                    return "Workbook" + i.ToString();

                i++;
            }
        }
        public object AddNew(string Objectname)
        {
            ExcelWorkbook retValue = new ExcelWorkbook();
            retValue.WorkbookName = New(Objectname);
            this.Add(retValue);
            return retValue;
        }
        public string New(string Objectname)
        {
            int i = 1;
            while (true)
            {
                if (this["Workbook" + i.ToString()] == null)
                    return "Workbook" + i.ToString();

                i++;
            }
        }
        public bool Rename(string OldObjectname, string NewObjectname)
        {
            ExcelWorkbook tProjectFlow = this.Where(f => f.ProjectItemID.Equals(OldObjectname)).FirstOrDefault();
            if (tProjectFlow != null)
            {
                tProjectFlow.ProjectItemName = NewObjectname;
                return true;
            }
            else
                return false;
        }
    }
}
