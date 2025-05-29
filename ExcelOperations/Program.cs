using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOperations
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelWorkbook wkBook = new ExcelWorkbook();
            wkBook.OpenExcelFile(Path.Combine(Environment.CurrentDirectory,"Input"));
            wkBook.ExcelApplication.Visible = true;
            wkBook.ExcelApplication.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;
            wkBook.ExcelApplication.Workbooks.Add(Missing.Value);
            string getValue = wkBook.GetCellValue("b6:b6",ExcelValueType.Value);           
            wkBook.InsertWorksheet("MyNewSheet", 2);
            wkBook.DeleteWorksheet(sheetName: "MyNewSheet");
            wkBook.ActiveSheet.Activate();
            wkBook.SetCellValue("a1:C3", "Value");
            wkBook.InsertColumnBefore("A:A");
            wkBook.InsertRowBefore(2);
            wkBook.DeleteRow("2:2");
            wkBook.DeleteColumn("B:B");
            wkBook.CloseWorkbook();
        }
    }
}
