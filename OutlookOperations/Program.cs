using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    internal class Program
    {
        static void Main(string[] args)
        {
            MSOutlookOperations.Instance.OpenOutlook();
            MSOutlookOperations.Instance.SendAndReceive(5);
            MSOutlookOperations.Instance.CloseOutlook();
            MSOutlookOperations.Instance.SendMail();
            MSOutlookOperations.Instance.ProcessMails("EmailBOXID");
        }
    }
   
}
