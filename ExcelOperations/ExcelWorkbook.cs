using Microsoft.Office.Interop.Excel;
using System;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using System.Collections.Generic;

namespace ExcelOperations
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

    public class ExcelWorkbook: NotifyProperyChangedBase,IComparable
    {
        private string workbookFileName;
        private string workbookname = string.Empty;
        private string mProjectItemName = string.Empty;
        private Application excelApplication;
        private int currentRow;
        private int currentCol;
        private string mProjectItemID = string.Empty;

        public ExcelWorkbook()
        {
            excelApplication = new Application();
        }
        public  string ProjectItemID
        {
            get { return mProjectItemID; }
            set
            {
                if (this.CheckPropertyChanged<string>("ProjectItemID", ref mProjectItemID, ref value))
                {
                    this.FirePropertyChanged("ProjectItemID");

                }
            }
        }
        public string WorkbookName
        {
            get
            {
                workbookname = ProjectItemName;
                return ProjectItemName;
            }
            set
            {
                if (this.CheckPropertyChanged<string>("WorkbookName", ref workbookname, ref value))
                {
                    ProjectItemName = workbookname;
                    this.FirePropertyChanged("WorkbookName");

                }
            }
        }
        public virtual string ProjectItemName
        {
            get { return mProjectItemName; }
            set
            {
                if (this.CheckPropertyChanged<string>("ProjectItemName", ref mProjectItemName, ref value))
                {
                    this.FirePropertyChanged("ProjectItemName");

                }
            }
        }
        public string WorkbookFileName
        {
            get { return workbookFileName; }
            set { workbookFileName = value; }
        }
        public Application ExcelApplication
        {
            get { return excelApplication; }
            set { excelApplication = value; }
        }
        public int CurrentWorksheetIndex
        {
            get
            {
                if (excelApplication == null) return -1;
                return ((Worksheet)excelApplication.ActiveSheet).Index;
            }
        }
        public string CurrentWorksheetName
        {
            get
            {
                if (excelApplication == null) return "";
                return ((Worksheet)excelApplication.ActiveSheet).Name;
            }
        }
        public int CurrentRow
        {
            get { return currentRow; }
            set { currentRow = value; }
        }
        public int CurrentColumn
        {
            get { return currentCol; }
            set { currentCol = value; }
        }
        public string CurrentColumnName
        {
            get { return ExcelTranslator.GetColNameFromNumber(currentCol); }
        }
        public int CurrentSelectionRow
        {
            get
            {
                if (excelApplication == null) return -1;
                return ((Range)excelApplication.Selection).Row;
            }
        }
        public int CurrentSelectionCol
        {
            get
            {
                if (excelApplication == null) return -1;
                return ((Range)excelApplication.Selection).Column;
            }
        }
        public string CurrentSelectionColName
        {
            get
            {
                if (excelApplication == null) return "";
                return ExcelTranslator.GetColNameFromNumber(((Range)excelApplication.Selection).Column);
            }
        }
        public string CurrentSelectionAddress
        {
            get
            {
                if (excelApplication == null) return "";
                return ((Range)excelApplication.Selection).get_Address(true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, Missing.Value, Missing.Value);
            }
        }
        public void OpenExcelFile(string Filename)
        {
            excelApplication.Visible = true;
            excelApplication.Workbooks.Open(Filename);
        }
        public void CloseWorkbook()
        {
            //wkBook.ExcelApplication.Workbooks.Close();
            if (ExcelApplication.Workbooks.Count > 0)
                ExcelApplication.Workbooks[1].Close(false, Missing.Value, Missing.Value);

            ExcelApplication.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApplication);
            ExcelApplication = null;
        }
        public string GetCellValue(int Row, int Col, ExcelValueType valType)
        {
            if (valType == ExcelValueType.Value)
            {
                object o = ((Range)((Worksheet)excelApplication.ActiveSheet).Cells[Row, Col]).Value2;
                return o == null ? "" : o.ToString();
            }
            else if (valType == ExcelValueType.Formula)
                return ((Range)((Worksheet)excelApplication.ActiveSheet).Cells[Row, Col]).Formula.ToString();
            else if (valType == ExcelValueType.Text)
                return ((Range)((Worksheet)excelApplication.ActiveSheet).Cells[Row, Col]).Text.ToString();
            return "";
        }
        public string GetCellValue(string Address, ExcelValueType valType)
        {
            if (valType == ExcelValueType.Value)
            {
                object o = ((Range)excelApplication.get_Range(Address, Missing.Value)).Value2;
                return o == null ? "" : o.ToString();
            }
            else if (valType == ExcelValueType.Formula)
                return ((Range)excelApplication.get_Range(Address, Missing.Value)).Formula.ToString();
            else if (valType == ExcelValueType.Text)
                return ((Range)excelApplication.get_Range(Address, Missing.Value)).Text.ToString();

            return "";
        }
        public void SetCellValue(int Row, int Col, object val)
        {
            ((Range)((Worksheet)excelApplication.ActiveSheet).Cells[Row, Col]).Value2 = val;
        }
        public void SetCellValue(string address, object val)
        {
            ((Range)excelApplication.get_Range(address, Missing.Value)).Value2 = val;
        }
        public void SetCellFormula(int Row, int Col, string val)
        {
            ((Range)((Worksheet)excelApplication.ActiveSheet).Cells[Row, Col]).Formula = val;
        }
        public void SetCellFormula(string address, string val)
        {
            ((Range)excelApplication.get_Range(address, Missing.Value)).Formula = val;
        }
        public ExcelCellFormat GetCellFormatting(int Row, int Col)
        {
            return new ExcelCellFormat(((Range)((Worksheet)excelApplication.ActiveSheet).Cells[Row, Col]));
        }
        public ExcelCellFormat GetCellFormatting(string address)
        {
            return new ExcelCellFormat(((Range)excelApplication.get_Range(address, Missing.Value)));
        }
        public void SetCellFormatting(int Row, int Col, ExcelCellFormat format)
        {
            format.SetFormatOnCell(((Range)((Worksheet)excelApplication.ActiveSheet).Cells[Row, Col]));
        }
        public void SetCellFormatting(string address, ExcelCellFormat format)
        {
            format.SetFormatOnCell(((Range)excelApplication.get_Range(address, Missing.Value)));
        }
        public void InsertRowBefore(int row)
        {
            this.Range(row.ToString() + ":" + row.ToString()).Rows.Insert(XlInsertShiftDirection.xlShiftDown, Missing.Value);
            this.Range(row.ToString() + ":" + row.ToString()).Rows.Select();
        }
        public void InsertRowBefore(string RowAddress)
        {
            this.Range(RowAddress.ToString() + ":" + RowAddress.ToString()).Rows.Insert(XlInsertShiftDirection.xlShiftDown, Missing.Value);
            this.Range(RowAddress.ToString() + ":" + RowAddress.ToString()).Rows.Select();
        }
        public Range Cells(int Row, int Col)
        {
            return (Range)excelApplication.Cells[Row, Col];
        }
        public Range Range(string Address)
        {
            return (Range)excelApplication.get_Range(Address, Missing.Value);
        }
        public Range UsedRange()
        {
            return (Range)((Worksheet)excelApplication.ActiveSheet).UsedRange;
        }
        public Range Range(string Address, object Type)
        {
            return (Range)excelApplication.get_Range(Address, Type);
        }
        public Range Range(string Address, object Type, object Value)
        {
            return (Range)excelApplication.get_Range(Address, Type);
        }
        public string UsedRangeAddress()
        {
            return ((Worksheet)excelApplication.ActiveSheet).UsedRange.get_Address(
                false, false, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, Missing.Value, Missing.Value);
        }
        public string Address(Range range)
        {
            return range.get_Address(false, false, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, Missing.Value, Missing.Value);
        }
        public System.Data.DataTable RangeFormulasToDataTable(string Address)
        {
            return RangeValuesToDataTable(Address);
        }
        public System.Data.DataTable RangeValuesToDataTable(string Address)
        {
            Range r = this.Range(Address);

            string address = r.get_Address(false, false, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, Missing.Value, Missing.Value);

            if (address.Contains(":")) address = address.Substring(0, address.IndexOf(":"));
            int StartCol = ExcelTranslator.GetColumnIndex(ExcelTranslator.GetColNameFromAddress(address));
            int StartRow = ExcelTranslator.GetRowNumber(address);
            if (StartRow < 1) StartRow = 1;


            object[,] values = (object[,])(r.get_Value(XlRangeValueDataType.xlRangeValueDefault));

            System.Data.DataTable valuesDT = new System.Data.DataTable();
            valuesDT.Columns.Add("RowNumber");

            for (int i = 0; i < values.GetLength(1); i++)
                valuesDT.Columns.Add(ExcelTranslator.GetColNameFromNumber(StartCol + i));

            for (int row = 0; row < values.GetLength(0); row++)
            {
                DataRow newRow = valuesDT.NewRow();

                newRow["RowNumber"] = StartRow + row;

                for (int i = 0; i < values.GetLength(1); i++)
                    newRow[ExcelTranslator.GetColNameFromNumber(StartCol + i)] = values[row + 1, i + 1];

                valuesDT.Rows.Add(newRow);
            }

            Marshal.ReleaseComObject(r);
            r = null;

            return valuesDT;
        }
        public void DataTableToRange(System.Data.DataTable source, string StartingCellAddress, bool outputHeaders)
        {
            if (!outputHeaders && source.Rows.Count == 0) return;

            object[,] test = new object[outputHeaders ? source.Rows.Count + 1 : source.Rows.Count, source.Columns.Count];

            int nextRowInput = 0;
            int nextRowOutput = 0;

            if (outputHeaders)
            {
                for (int j = 0; j < source.Columns.Count; j++)
                {
                    test[nextRowOutput, j] = source.Columns[j].ColumnName;
                }

                nextRowOutput++;
            }

            for (int i = 0; i < source.Rows.Count; i++)
            {
                for (int j = 0; j < source.Columns.Count; j++)
                {
                    test[nextRowOutput, j] = source.Rows[i][j];
                }

                nextRowOutput++;
            }

            Range r = ((Worksheet)excelApplication.ActiveSheet).get_Range(StartingCellAddress, Missing.Value);
            r = r.get_Resize(test.GetLength(0), test.GetLength(1));
            r.set_Value(XlRangeValueDataType.xlRangeValueDefault, test);
            Marshal.ReleaseComObject(r);
            r = null;
        }
        public static object GetExcelDataTableValue(System.Data.DataTable table, string address)
        {
            int Row = ExcelTranslator.GetRowNumber(address);
            string Col = ExcelTranslator.GetColNameFromAddress(address);
            DataRow[] find = table.Select("RowNumber = " + Row);
            if (find.Length == 0) throw new Exception("Excel Row not found in data table");
            if (!find[0].Table.Columns.Contains(Col)) throw new Exception("Excel Column not found in data table");
            return find[0][Col];
        }
        public static object GetExcelDataTableValue(System.Data.DataTable table, int Row, int Col)
        {
            return GetExcelDataTableValue(table, ExcelTranslator.GetColNameFromNumber(Col) + Row.ToString());
        }
        public static void SetExcelDataTableValue(System.Data.DataTable table, string address, object value)
        {
            int Row = ExcelTranslator.GetRowNumber(address);
            string Col = ExcelTranslator.GetColNameFromAddress(address);
            DataRow[] find = table.Select("RowNumber = " + Row);
            if (find.Length == 0) throw new Exception("Excel Row not found in data table");
            if (!find[0].Table.Columns.Contains(Col)) throw new Exception("Excel Column not found in data table");
            find[0][Col] = value;
        }
        public static void SetExcelDataTableValue(System.Data.DataTable table, int Row, int Col, object value)
        {
            SetExcelDataTableValue(table, ExcelTranslator.GetColNameFromNumber(Col) + Row.ToString(), value);
        }
        public bool FindAndActivate(string RangeAddress, object ValueToFind, ExcelDirection SearchDirection, ExcelValueType SearchType, bool MatchEntireContents)
        {
            Range r = this.excelApplication.get_Range(RangeAddress, Missing.Value);
            bool result = FindAndActivate(r, ValueToFind, SearchDirection, SearchType, MatchEntireContents);
            Marshal.ReleaseComObject(r);
            r = null;
            return result;
        }
        public bool FindAndActivate(Range Range, object ValueToFind, ExcelDirection SearchDirection, ExcelValueType SearchType, bool MatchEntireContents)
        {
            bool resultVal = false;
            Range result = null;

            try
            {
                result = Range.Find(ValueToFind, Missing.Value,
                    SearchType == ExcelValueType.Formula ? XlFindLookIn.xlFormulas : XlFindLookIn.xlValues,
                    MatchEntireContents ? XlLookAt.xlWhole : XlLookAt.xlPart,
                    SearchDirection == ExcelDirection.Right || SearchDirection == ExcelDirection.Left ? XlSearchOrder.xlByRows : XlSearchOrder.xlByColumns,
                    SearchDirection == ExcelDirection.Right || SearchDirection == ExcelDirection.Down ? XlSearchDirection.xlNext : XlSearchDirection.xlPrevious,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value);
            }
            catch (Exception ex)
            {
                //a type mismatch error means that it didn't find
                //any data types in Excel to compare the value to and
                //it freaks it out.  So we're going to count this as a non-match                
                if (!ex.Message.Contains("Type mismatch")) throw ex;
            }

            if (result != null)
            {
                resultVal = true;
                result.Select();

                Marshal.ReleaseComObject(result);
                result = null;
            }

            return resultVal;
        }
        public object GetFormControlProperty(string ControlName, string PropertyName, object PropertyParameter)
        {
            try
            {
                if (this.ActiveSheet.Shapes.Item(ControlName).Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                {
                    if (PropertyName == "Text")
                        return ActiveSheet.Shapes.Item(ControlName).TextFrame.Characters(Missing.Value, Missing.Value).Text;
                }
            }
            catch
            {
            }

            object formControl = GetVariable(this.ActiveSheet, ControlName);
            object propertyValue = null;

            if (PropertyParameter == null)
                propertyValue =  GetVariable(formControl, PropertyName);
            else
            {
                object[] toSend = new object[1];
                toSend[0] = PropertyParameter;
                propertyValue = GetVariable(formControl, PropertyName, toSend);
            }

            return propertyValue;
        }
        public void SetFormControlProperty(string ControlName, string PropertyName, object ValueToSet)
        {
            try
            {
                if (this.ActiveSheet.Shapes.Item(ControlName).Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                {
                    if (PropertyName == "Text")
                    {
                        ActiveSheet.Shapes.Item(ControlName).TextFrame.Characters(Missing.Value, Missing.Value).Text = ValueToSet.ToString();
                        return;
                    }
                }
            }
            catch
            {
            }

            object formControl = GetVariable(this.ActiveSheet, ControlName);
            SetVariable(formControl, PropertyName, ValueToSet);
        }
        public int CompareTo(object obj)
        {
            if (obj is ExcelWorkbook)
                return this.WorkbookName.CompareTo(((ExcelWorkbook)obj).WorkbookName);
            else
                return 0;
        }
        public Worksheet ActiveSheet
        {
            get
            {
                return (Worksheet)excelApplication.ActiveSheet;
            }
        }     
        public static object GetVariable(object obj, string VariableName)
        {
            return GetVariable(obj, VariableName, new object[] { });
        }
        public static object GetVariable(object obj, string VariableName, object[] param)
        {
            Type type = obj.GetType();

            return type.InvokeMember(VariableName, BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetField | BindingFlags.GetProperty,
                null, obj, param);
        }
        public static void SetVariable(object obj, string VariableName, object value)
        {
            object[] valueArray = new object[1];
            valueArray[0] = value;

            Type type = obj.GetType();

            type.InvokeMember(VariableName,
                BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetField | BindingFlags.SetProperty,
                null, obj, valueArray);
        }

        public void DeleteWorksheet(string sheetName)
        {
            excelApplication.DisplayAlerts = false;
            Worksheet worksheet = (Worksheet)excelApplication.Worksheets[sheetName];
            worksheet.Delete();
            excelApplication.DisplayAlerts = true;
        }
        public void InsertWorksheet(string sheetName,int SheetsCountInsertAfter= 0,int SheetsCountInsertBefore = 0)
        {
            Worksheet objWorksheet= (Worksheet)excelApplication.Sheets.Add(Missing.Value,Missing.Value, Missing.Value, Missing.Value);
            if (SheetsCountInsertAfter != 0)
            {
                 objWorksheet = (Worksheet)excelApplication.Sheets.Add(Missing.Value, excelApplication.Sheets[SheetsCountInsertAfter], Missing.Value, Missing.Value);
            }
            else if (SheetsCountInsertBefore != 0)
            {
                 objWorksheet = (Worksheet)excelApplication.Sheets.Add(excelApplication.Sheets[SheetsCountInsertBefore], Missing.Value, Missing.Value, Missing.Value);
            }

            objWorksheet.Name = sheetName;

        }
        public void DeleteRow(string row)
        {
            //this.Range(row.ToString() + ":" + row.ToString()).Rows.Delete(XlDeleteShiftDirection.xlShiftUp);
            this.Range(row).Rows.Delete(XlDeleteShiftDirection.xlShiftUp);
        }
        public void InsertColumnBefore(int Col)
        {
            string addr = ExcelTranslator.GetColNameFromNumber(Col);
            addr = addr + ":" + addr;
            this.Range(addr).Columns.Insert(XlInsertShiftDirection.xlShiftToRight, Missing.Value);
            this.Range(addr).Columns.Select();
        }
        public void InsertColumnBefore(string ColAddress)
        {
            this.Range(ColAddress).Columns.Insert(XlInsertShiftDirection.xlShiftToRight, Missing.Value);
            this.Range(ColAddress).Columns.Select();
        }
        public void DeleteColumn(int Col)
        {



            string addr = ExcelTranslator.GetColNameFromNumber(Col);
            addr = addr + ":" + addr;
            this.Range(addr).Activate();
        }

            public void DeleteColumn(string ColAddress)
        {

            //string addr = ExcelTranslator.GetColNameFromNumber(Col);
            //addr = addr + ":" + addr;
            string tempAct = ColAddress.Split(':')[0].ToString();
            this.Range(tempAct + "2").Activate();
            this.Range(ColAddress).EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
        }

        #region Create a PivotTable  
        public PivotTable InsertPivotTable(string DestinationSheetName, string SourceSheetname, string SourceData)
        {

            Workbook book = this.excelApplication.ActiveWorkbook;
            Worksheet objWorksheet = (Worksheet)excelApplication.Sheets[DestinationSheetName];
            PivotCaches pCaches = book.PivotCaches();
            PivotCache pCache = pCaches.Create(XlPivotTableSourceType.xlDatabase, SourceData, XlPivotTableVersionList.xlPivotTableVersion14);
            Range rngDes = objWorksheet.get_Range("A1");
            PivotTable pTable = pCache.CreatePivotTable(TableDestination: rngDes, TableName: "PivotTable1", DefaultVersion: XlPivotTableVersionList.xlPivotTableVersion14);
            return pTable;
        }
        public void SetPivotDataField(PivotTable pivotTable, string FieldName, string FunctionName)
        {
            PivotField fAmount = pivotTable.PivotFields(FieldName);
            fAmount.Orientation = XlPivotFieldOrientation.xlDataField;
            string[] arrValue = new string[] { "of" };
            string[] strValue = FunctionName.Split(arrValue, 10, StringSplitOptions.RemoveEmptyEntries);
            if (strValue.Length != 0)
            {
                FunctionName = strValue[0].ToString();
            }

            switch (FunctionName)
            {
                case "Sum":
                    fAmount.Function = XlConsolidationFunction.xlSum;
                    break;
                case "Count":
                    fAmount.Function = XlConsolidationFunction.xlCount;
                    break;
                case "Average":
                    fAmount.Function = XlConsolidationFunction.xlAverage;
                    break;
                case "Min":
                    fAmount.Function = XlConsolidationFunction.xlMin;
                    break;
                case "Max":
                    fAmount.Function = XlConsolidationFunction.xlMax;
                    break;
                case "Product":
                    fAmount.Function = XlConsolidationFunction.xlProduct;
                    break;
            }
        }
        public void SetPivotRowField(PivotTable pivotTable, string FieldName)
        {
            PivotField fQ = pivotTable.PivotFields(FieldName);
            fQ.Orientation = XlPivotFieldOrientation.xlRowField;
        }
        public void SetPivotColumnField(PivotTable pivotTable, string FieldName)
        {
            PivotField fQ = pivotTable.PivotFields(FieldName);
            fQ.Orientation = XlPivotFieldOrientation.xlColumnField;
        }

        #endregion
    }

    public class ExcelCellFormat
    {
        public ExcelCellFormat(Range cell)
        {
            FontStyle thisFontStyle = FontStyle.Regular;
            if ((bool)cell.Font.Bold) thisFontStyle = thisFontStyle | FontStyle.Bold;
            if ((bool)cell.Font.Italic) thisFontStyle = thisFontStyle | FontStyle.Italic;
            if ((XlUnderlineStyle)(cell.Font.Underline) != XlUnderlineStyle.xlUnderlineStyleNone) thisFontStyle = thisFontStyle | FontStyle.Underline;
            if ((bool)cell.Font.Strikethrough) thisFontStyle = thisFontStyle | FontStyle.Strikeout;

            //this.cellFont = new Microsoft.Office.Interop.Excel.Font(cell.Font.Name.ToString(), float.Parse(cell.Font.Size.ToString()), thisFontStyle);
            this.foregroundColor = ColorTranslator.FromOle(int.Parse(cell.Font.Color.ToString()));
            this.backgroundColor = ColorTranslator.FromOle(int.Parse(cell.Interior.Color.ToString()));

            this.borders = System.Windows.Forms.AnchorStyles.None;
            if ((int)cell.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle > 0)
                this.borders = this.borders | System.Windows.Forms.AnchorStyles.Top;

            if ((int)cell.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle > 0)
                this.borders = this.borders | System.Windows.Forms.AnchorStyles.Left;

            if ((int)cell.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle > 0)
                this.borders = this.borders | System.Windows.Forms.AnchorStyles.Right;

            if ((int)cell.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle > 0)
                this.borders = this.borders | System.Windows.Forms.AnchorStyles.Bottom;
        }
        public void SetFormatOnCell(Range cell)
        {
            cell.Font.Bold = CellFont.Bold;
            cell.Font.Italic = CellFont.Italic;
            cell.Font.Underline = CellFont.Underline;
            cell.Font.Strikethrough = CellFont.Strikethrough;

            cell.Font.Name = CellFont.Name;
            cell.Font.Color = ColorTranslator.ToOle(foregroundColor);
            cell.Interior.Color = ColorTranslator.ToOle(backgroundColor);

            cell.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                this.Borders == System.Windows.Forms.AnchorStyles.Top ? 1 : -4142;

            cell.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                this.Borders == System.Windows.Forms.AnchorStyles.Left ? 1 : -4142;

            cell.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                this.Borders == System.Windows.Forms.AnchorStyles.Right ? 1 : -4142;

            cell.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                this.Borders == System.Windows.Forms.AnchorStyles.Bottom ? 1 : -4142;

        }

        private Microsoft.Office.Interop.Excel.Font cellFont;
        public Microsoft.Office.Interop.Excel.Font CellFont
        {
            get { return cellFont; }
            set { cellFont = value; }
        }

        private Color backgroundColor = Color.Transparent;
        public Color BackgroundColor
        {
            get { return backgroundColor; }
            set { backgroundColor = value; }
        }

        private Color foregroundColor = Color.Black;
        public Color ForegroundColor
        {
            get { return foregroundColor; }
            set { foregroundColor = value; }
        }

        private System.Windows.Forms.AnchorStyles borders = System.Windows.Forms.AnchorStyles.None;
        public System.Windows.Forms.AnchorStyles Borders
        {
            get { return borders; }
            set { borders = value; }
        }
    }

    public class CreatePivotTable
    {
        public enum PivotFunction
        {
            None,
            Sum,
            Count,
            Average,
        }

        public string SheetName;
        private System.Data.DataTable sampleData;
        private System.Data.DataTable MappingDataTable;

        private autoDataTable vmsDatatable;
        private List<PivotTableMapping> mapping;
        private string firstColumnVar = "A";
        private string lastColumnVar = "G";
        private bool Loading = true;
        public static string LastExcelFileToGetSample = string.Empty;

        public object StartCol { get; private set; }
        public object EndCol { get; private set; }

        public string FirstColumn
        {
            get { return firstColumnVar; }
            set { firstColumnVar = value; }
        }
        public string LastColumn
        {
            get { return lastColumnVar; }
            set { lastColumnVar = value; }
        }

        public CreatePivotTable()
        {

        }

        public void ReloadExcelColumns()
        {
            System.Data.DataTable old = this.MappingDataTable.Copy();
            MappingDataTable.Rows.Clear();

            int startColIndex = ExcelTranslator.GetColumnIndex(this.StartCol.ToString() == "" ? "A" : this.StartCol.ToString());
            int endColIndex = ExcelTranslator.GetColumnIndex(this.EndCol.ToString() == "" ? "E" : this.EndCol.ToString());

            for (int i = startColIndex; i <= endColIndex; i++)
            {
                DataRow row = MappingDataTable.NewRow();
                row["ExcelColumnDisplay"] = ExcelTranslator.GetColNameFromNumber(i);


                DataRow[] search = old.Select("ExcelColumnDisplay= '" + ExcelTranslator.GetColNameFromNumber(i) + "'");
                if (search.Length > 0)
                {
                    row["PivotTableFields"] = search[0]["PivotTableFields"];
                    row["PivotTableColumns"] = search[0]["PivotTableColumns"];
                    row["PivotTableRows"] = search[0]["PivotTableRows"];
                    row["PivotTableValues"] = search[0]["PivotTableValues"];
                    row["ExcelColumnDisplay"] = search[0]["ExcelColumnDisplay"];
                }
                else
                {
                    row["ExcelColumnDisplay"] = ExcelTranslator.GetColNameFromNumber(i);
                }

                if (sampleData != null && sampleData.Rows.Count > 0)
                {
                    string temp = (sampleData.Rows[0][ExcelTranslator.GetColNameFromNumber(i)]).ToString();
                    if (sampleData != null && sampleData.Rows.Count > 0)
                    {
                        row["PivotTableFields"] = (sampleData.Rows[0][ExcelTranslator.GetColNameFromNumber(i)]).ToString();
                    }
                }
                if (row["PivotTableFields"].ToString() != string.Empty)
                {
                    MappingDataTable.Rows.Add(row);
                }
            }
        }

        public CreatePivotTable(List<PivotTableMapping> ptMapping, autoDataTable vwDataTable,
          string FirstColumn, string LastColumn)
        {

            Loading = true;

            this.vmsDatatable = vwDataTable;

            if (vwDataTable == null) return;

            List<string> cols = new List<string>();
            cols.Add("<Ignore>");
            cols.Add("Sum");
            cols.Add("Count");
            cols.Add("Average");
            cols.Add("Min");
            cols.Add("Max");
            cols.Add("Product");

            string Title = " [ Datatable Name :" + vwDataTable.DataTableName + " ]";
            this.mapping = ptMapping;

            this.FirstColumn = FirstColumn;
            this.LastColumn = LastColumn;

            this.StartCol = this.FirstColumn;
            this.EndCol = this.LastColumn;

            MappingDataTable = new System.Data.DataTable();
            MappingDataTable.Columns.Add("PivotTableFields", typeof(string));
            MappingDataTable.Columns.Add("PivotTableColumns", typeof(string));
            MappingDataTable.Columns.Add("PivotTableRows", typeof(string));
            MappingDataTable.Columns.Add("PivotTableValues", typeof(string));
            MappingDataTable.Columns.Add("ExcelColumnDisplay", typeof(string));

            foreach (PivotTableMapping map in mapping)
            {
                if (map.PivotTableFields != string.Empty)
                {
                    DataRow row = MappingDataTable.NewRow();
                    row["PivotTableFields"] = map.PivotTableFields;

                    if (map.PivotTableColumns.Equals("True") || map.PivotTableColumns.Equals("False"))
                        row["PivotTableColumns"] = false;
                    else
                        row["PivotTableColumns"] = true;

                    if (map.PivotTableRows.Equals("True") || map.PivotTableRows.Equals("False"))
                        row["PivotTableRows"] = false;
                    else
                        row["PivotTableRows"] = true;

                    if (map.PivotTableValues != string.Empty)
                    {
                        string[] arr1 = new string[] { "of" };
                        string[] str1 = map.PivotTableValues.Split(arr1, 10, StringSplitOptions.RemoveEmptyEntries);
                        if (str1.Length != 0)
                        {
                            row["PivotTableValues"] = str1[0].ToString().Trim();
                        }
                    }
                    else
                        row["PivotTableValues"] = map.PivotTableValues;

                    row["ExcelColumnDisplay"] = map.ExcelColumnDisplay;

                    MappingDataTable.Rows.Add(row);
                }
            }
            //  gridPivotTable.BindDataTableDataGrid(MappingDataTable);
            ReloadExcelColumns();
            Loading = false;
        }

        public CreatePivotTable(List<PivotTableMapping> ptMapping, autoDataTable vwDataTable)
        {
            Loading = true;

            this.vmsDatatable = vwDataTable;

            if (vwDataTable == null) return;

            List<string> cols = new List<string>();
            cols.Add("<Ignore>");
            cols.Add("Sum");
            cols.Add("Count");
            cols.Add("Average");
            cols.Add("Min");
            cols.Add("Max");
            cols.Add("Product");

            string Title = " [ Datatable Name :" + vwDataTable.DataTableName + " ]";
            this.mapping = ptMapping;

            MappingDataTable = new System.Data.DataTable();
            MappingDataTable.Columns.Add("PivotTableFields", typeof(string));
            MappingDataTable.Columns.Add("PivotTableColumns", typeof(string));
            MappingDataTable.Columns.Add("PivotTableRows", typeof(string));
            MappingDataTable.Columns.Add("PivotTableValues", typeof(string));
            MappingDataTable.Columns.Add("ExcelColumnDisplay", typeof(string));

            foreach (PivotTableMapping map in mapping)
            {
                if (map.PivotTableFields != string.Empty)
                {
                    DataRow row = MappingDataTable.NewRow();
                    row["PivotTableFields"] = map.PivotTableFields;

                    if (map.PivotTableColumns.Equals("True") || map.PivotTableColumns.Equals("False"))
                        row["PivotTableColumns"] = false;
                    else
                        row["PivotTableColumns"] = true;

                    if (map.PivotTableRows.Equals("True") || map.PivotTableRows.Equals("False"))
                        row["PivotTableRows"] = false;
                    else
                        row["PivotTableRows"] = true;

                    if (map.PivotTableValues != string.Empty)
                    {
                        string[] arr1 = new string[] { "of" };
                        string[] str1 = map.PivotTableValues.Split(arr1, 10, StringSplitOptions.RemoveEmptyEntries);
                        if (str1.Length != 0)
                        {
                            row["PivotTableValues"] = str1[0].ToString().Trim();
                        }
                    }
                    else
                        row["PivotTableValues"] = map.PivotTableValues;

                    row["ExcelColumnDisplay"] = map.ExcelColumnDisplay;

                    MappingDataTable.Rows.Add(row);
                }
            }
            //  gridPivotTable.BindDataTableDataGrid(MappingDataTable);
            ReloadExcelColumns();
        }


    }

    public class PivotTableMapping
    {
        private string pivotTableFields = string.Empty;
        private string pivotTableColumns = string.Empty;
        private string pivotTableRows = string.Empty;
        private string pivotTableValues = string.Empty;
        private string excelColumnDisplay = string.Empty;


        public PivotTableMapping()
        {

        }

        public PivotTableMapping(string ptFields, string ptColumns, string ptRows, string ptValues, string excelColumn)
        {
            pivotTableFields = ptFields;
            pivotTableColumns = ptColumns;
            pivotTableRows = ptRows;
            pivotTableValues = ptValues;
            excelColumnDisplay = excelColumn;
        }

        public string ExcelColumnDisplay
        {
            get { return excelColumnDisplay; }
            set { excelColumnDisplay = value; }
        }

        public string PivotTableValues
        {
            get { return pivotTableValues; }
            set { pivotTableValues = value; }
        }

        public string PivotTableRows
        {
            get { return pivotTableRows; }
            set { pivotTableRows = value; }
        }

        public string PivotTableColumns
        {
            get { return pivotTableColumns; }
            set { pivotTableColumns = value; }
        }
        public string PivotTableFields
        {
            get { return pivotTableFields; }
            set { pivotTableFields = value; }
        }
    }

    public class autoDataTable : IComparable
    {
        public autoDataTable()
        {
        }

        public autoDataTable(string DataTableName)
        {
            this.DataTableName = DataTableName;
        }

        public autoDataTable(string DataTableName, System.Data.DataTable table)
        {
            this.DataTableName = DataTableName;
            this.DataTable = table;
        }

        private string mDataTableName;

        public string DataTableName
        {
            get { return mDataTableName; }
            set { mDataTableName = value; }
        }

        private Columns mColumns = new Columns();

        public Columns Columns
        {
            get { return mColumns; }
            set { mColumns = value; }
        }

        [XmlIgnore]
        public System.Data.DataTable DataTable = new System.Data.DataTable();

        [XmlIgnore]
        public int CurrentRowIndex = -1;

        [XmlIgnore]
        public DataRowCollection Rows
        {
            get
            {
                if (DataTable == null) return null;
                return DataTable.Rows;
            }
        }

        [XmlIgnore]
        public DataRow CurrentRow
        {
            get
            {
                if (DataTable == null) return null;

                if (CurrentRowIndex > -1 && CurrentRowIndex < DataTable.Rows.Count)
                    return DataTable.Rows[CurrentRowIndex];

                return null;
            }
            set
            {
                if (DataTable == null) return;
                if (DataTable.Rows.IndexOf(value) != -1)
                    CurrentRowIndex = DataTable.Rows.IndexOf(value);
            }
        }



        public bool BOF
        {
            get
            {
                if (DataTable == null) return true;
                return (CurrentRowIndex < 0);
            }
        }

        public bool EOF
        {
            get
            {
                if (DataTable == null) return true;
                return (CurrentRowIndex >= DataTable.Rows.Count);
            }
        }

        public bool NextRow()
        {
            CurrentRowIndex++;
            if (DataTable == null) return false;
            return (CurrentRowIndex < DataTable.Rows.Count);
        }

        public bool FirstRow()
        {
            CurrentRowIndex = 0;
            if (DataTable == null) return false;
            return (CurrentRowIndex < DataTable.Rows.Count);
        }

        public bool LastRow()
        {
            if (DataTable == null) return false;
            CurrentRowIndex = DataTable.Rows.Count - 1;
            return (CurrentRowIndex < DataTable.Rows.Count);
        }

        public bool PreviousRow()
        {
            CurrentRowIndex--;
            if (DataTable == null) return false;
            return (CurrentRowIndex >= 0) && (CurrentRowIndex < DataTable.Rows.Count);
        }

        public object GetFieldValue(string FieldName)
        {
            return CurrentRow[FieldName];
        }

        public void SetFieldValue(string FieldName, object FieldValue)
        {
            CurrentRow[FieldName] = FieldValue;
        }

        public void CreateBlankDataTableWithColumns()
        {
            DataTable = new System.Data.DataTable();
            foreach (Column col in Columns)
            {
                if (string.IsNullOrEmpty(col.TypeName)) col.TypeName = "string";

                Type t = Type.GetType(col.TypeName);
                if (t == null)
                {
                    if (col.TypeName.ToLower() == "string") t = typeof(string);
                    if (col.TypeName.ToLower() == "int") t = typeof(int);
                    if (col.TypeName.ToLower() == "datetime") t = typeof(DateTime);
                    if (col.TypeName.ToLower() == "decimal") t = typeof(decimal);
                    if (col.TypeName.ToLower() == "bool") t = typeof(bool);

                    if (t == null) t = typeof(string);
                }

                DataTable.Columns.Add(col.Name, t);
            }
        }

        public int CompareTo(object obj)
        {
            if (obj is autoDataTable)
                return this.DataTableName.CompareTo(((autoDataTable)obj).DataTableName);
            else
                return 0;
        }



    }

    public class Column
    {
        private string mName = string.Empty;
        private string mTypeName = string.Empty;
        private string mColumnID = string.Empty;

        public string ColumnID
        {
            get { return mColumnID; }
            set { mColumnID = value; }
        }
        public string Name
        {
            get { return mName; }
            set { mName = value; }
        }

        public string TypeName
        {
            get { return mTypeName; }
            set { mTypeName = value; }
        }

        public Column()
        {
        }

        public Column(string Name)
        {
            this.Name = Name;
        }
    }

    public class Columns : List<Column>
    {
        public Column this[string ColumnName]
        {
            get
            {
                foreach (Column c in this)
                {
                    if (c.Name.Equals(ColumnName, StringComparison.CurrentCultureIgnoreCase)) return c;
                }

                return null;
            }
            set
            {
                this[ColumnName] = value;
            }
        }

        public int IndexOf(string ColumnName)
        {
            if (this[ColumnName] == null) return -1;
            return this.IndexOf(this[ColumnName]);
        }

        public bool Contains(string ColumnName)
        {
            return (!(this[ColumnName] == null));
        }

        public Column GetByColumnName(string ColumnName)
        {
            return this[ColumnName];
        }

        public Column AddNewColumn()
        {
            Column retValue = new Column();
            retValue.Name = GetNextColumnName();
            this.Add(retValue);
            return retValue;
        }

        public string GetNextColumnName()
        {
            int i = 1;
            while (true)
            {
                if (this["Column" + i.ToString()] == null)
                    return "Column" + i.ToString();

                i++;
            }
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
            Column tProjectFlow = this.Where(f => f.Name.Equals(Objectname)).FirstOrDefault();
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
            Column retValue = new Column();
            retValue.Name = New();
            this.Add(retValue);
            return retValue;
        }

        public string New()
        {
            int i = 1;
            while (true)
            {
                if (this["Column" + i.ToString()] == null)
                    return "Column" + i.ToString();

                i++;
            }
        }

        public bool Rename(string OldObjectname, string NewObjectname)
        {
            Column tProjectFlow = this.Where(f => f.Name.Equals(OldObjectname)).FirstOrDefault();
            if (tProjectFlow != null)
            {
                tProjectFlow.Name = NewObjectname;
                return true;
            }
            else
                return false;
        }
    }

    public class ExcelTranslator
    {
        public static string GetColNameFromNumber(int number)
        {
            string ret = "";
            while (number > 0)
            {
                ret = (char)(--number % 26 + 'A') + ret;
                number /= 26;


            }
            return ret;
        }
        public static int GetColumnIndex(string colName)
        {
            colName = colName.ToUpper();
            int colNumber = 0;

            while (colName != "")
            {
                if ((int)colName[0] < 65 || (int)colName[0] > 90) return -1; //invalid char

                colNumber += (int)Math.Pow(26, (colName.Length - 1)) * ((int)colName[0] - 64);
                if (colName.Length == 1)
                    colName = "";
                else
                    colName = colName.Substring(1);
            }

            return colNumber;
        }
        public static int GetRowNumber(string cellAddress)
        {
            if (cellAddress.Length < 2) return -1; //invalid address
            if ((int)cellAddress[0] < 65 || (int)cellAddress[0] > 90) return -1; //invalid address

            for (int i = 1; i < cellAddress.Length; i++)
            {
                int temp;
                if (int.TryParse(cellAddress[i].ToString(), out temp))
                {
                    if (int.TryParse(cellAddress.Substring(i), out temp))
                        return temp;
                    else
                        return -1;
                }
                else
                {
                    if ((int)cellAddress[i] < 65 || (int)cellAddress[i] > 90) return -1; //invalid address
                }
            }

            return -1;
        }
        public static string GetColNameFromAddress(string cellAddress)
        {
            cellAddress = cellAddress.ToUpper();

            if ((int)cellAddress[0] < 65 || (int)cellAddress[0] > 90) return "";
            if (cellAddress.Length < 2)
            {
                if (cellAddress.Length == 1)
                    return cellAddress[0].ToString();

                return "";
            }


            for (int i = 1; i < cellAddress.Length; i++)
            {
                int temp;
                if (int.TryParse(cellAddress[i].ToString(), out temp))
                {
                    if (int.TryParse(cellAddress.Substring(i), out temp))
                        return cellAddress.Substring(0, i);
                    else
                        return "";
                }
                else
                {
                    if ((int)cellAddress[i] < 65 || (int)cellAddress[i] > 90) return ""; //invalid address
                }
            }

            return "";
        }
    }

    #region Exceldata
    [Serializable]
    public enum ExcelValueType
    {
        Value, Formula, Text
    }

    [Serializable]
    public enum ExcelDirection
    {
        None, Up, Down, Left, Right
    }
    [Serializable]
    public enum ExcelTextFileType
    {
        Delimited, FixedWidth
    }
    [Serializable]
    public enum ExcelTextFileTextQualifier
    {
        None, SingleQuote, DoubleQuote
    }  
  
    #endregion

   
}