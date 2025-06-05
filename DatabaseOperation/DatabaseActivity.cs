using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace DatabaseOperation
{
    public class DatabaseActivity
    {
        public DatabaseActivity()
        {

        }

        public static void DatabaseActivitys(string ConnectionName, string ConnectionString, string ProviderType, string QueryText)
        {
            object queryScalarResult = null;
            DbConnection connection = null;
            DbCommand command = null;
            if (!string.IsNullOrEmpty(ProviderType))
                connection = DataProviders.CreateDatabaseConnection(ProviderType, ConnectionString);
            else
                connection = DataProviders.CreateDatabaseConnection("System.Data.SqlClient", ConnectionString);


            if (connection != null)
            {
                try
                {
                    command = DataProviders.CreateCommand(QueryText, connection);
                    connection.Open();
                    queryScalarResult = command.ExecuteScalar();

                }
                finally
                {
                    if (connection.State == ConnectionState.Open)
                        connection.Close();
                }
            }
        }

    }

    public class DatabaseConnection : IComparable
    {
        private string mConnectionName = string.Empty;
        private string mProviderType = string.Empty;
        private string mConnectionString = string.Empty;
        public string ProjectItemID
        {
            get; set;
        }
        public string ProjectItemName
        {
            get; set;
        }
        public string ConnectionName
        {
            get;
            set;
        }

        public string ConnectionString
        {
            get;
            set;
        }
        public string ProviderType
        {
            get;
            set;
        }

        public override string ToString()
        {
            return ConnectionString;
        }

        public int CompareTo(object obj)
        {
            if (obj is DatabaseConnection)
                return this.ConnectionName.CompareTo(((DatabaseConnection)obj).ConnectionName);
            else
                return 0;
        }
    }

    public class DatabaseConnections : List<DatabaseConnection>
    {
        public DatabaseConnection this[string ConnectionName]
        {
            get
            {
                foreach (DatabaseConnection d in this)
                    if (d.ConnectionName.Equals(ConnectionName, StringComparison.CurrentCultureIgnoreCase)) return d;

                return null;
            }
        }

        public bool Delete(string Objectname)
        {

            DatabaseConnection tProjectFlow = this.Where(f => f.ProjectItemID.Equals(Objectname)).FirstOrDefault();
            if (tProjectFlow != null)
            {
                return this.Remove(tProjectFlow);
            }
            else
                return false;
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
            DatabaseConnection retValue = new DatabaseConnection();
            retValue.ConnectionName = New();
            this.Add(retValue);
            return retValue;
        }

        public string New()
        {
            int i = 1;
            while (true)
            {
                if (this["DBConnection" + i.ToString()] == null)
                    return "DBConnection" + i.ToString();

                i++;
            }
        }

        public bool Rename(string OldObjectname, string NewObjectname)
        {
            DatabaseConnection tProjectFlow = this.Where(f => f.ProjectItemID.Equals(OldObjectname)).FirstOrDefault();
            if (tProjectFlow != null)
            {
                tProjectFlow.ProjectItemName = NewObjectname;
                return true;
            }
            else
                return false;
        }
    }

    public class AutomationDataTable : IComparable
    {
        public AutomationDataTable()
        {
        }

        public AutomationDataTable(string DataTableName)
        {
            this.DataTableName = DataTableName;
        }
        public string ProjectItemID
        {
            get; set;
        }
        public string ProjectItemName
        {
            get; set;
        }

        public AutomationDataTable(string DataTableName, DataTable table)
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
        public DataTable DataTable = new DataTable();

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
            DataTable = new DataTable();
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
            if (obj is AutomationDataTable)
                return this.DataTableName.CompareTo(((AutomationDataTable)obj).DataTableName);
            else
                return 0;
        }



    }

    public class AutomationDataTables : List<AutomationDataTable>
    {
        public AutomationDataTables()
        {
        }

        public AutomationDataTable this[string index]
        {
            get
            {
                foreach (AutomationDataTable def in this)
                    if (def.DataTableName.Equals(index, StringComparison.CurrentCultureIgnoreCase)) return def;

                return null;
            }
        }

        public AutomationDataTable Table(string TableName)
        {
            return this[TableName];
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
            AutomationDataTable tVWSDataTable = this.Where(f => f.ProjectItemID.Equals(Objectname)).FirstOrDefault();
            if (tVWSDataTable != null)
            {
                return this.Remove(tVWSDataTable);
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
            AutomationDataTable retValue = new AutomationDataTable();
            retValue.DataTableName = New();
            this.Add(retValue);
            return retValue;
        }

        public string New()
        {
            int i = 1;
            while (true)
            {
                if (this["DataTable" + i.ToString()] == null)
                    return "DataTable" + i.ToString();

                i++;
            }
        }

        public bool Rename(string OldObjectname, string NewObjectname)
        {
            AutomationDataTable tProjectFlow = this.Where(f => f.ProjectItemID.Equals(OldObjectname)).FirstOrDefault();
            if (tProjectFlow != null)
            {
                tProjectFlow.ProjectItemName = NewObjectname;
                return true;
                ;

            }
            else
                return false;
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

    public class DataProviders
    {
        private static DbProviderFactory dbProviderFactory;
        public static DbConnection CreateDatabaseConnection(string ProviderName, string ConnectionString)
        {
            DbConnection connection = null;
            dbProviderFactory = null;
            if (string.IsNullOrEmpty(ConnectionString) || string.IsNullOrEmpty(ProviderName)) return null;

            dbProviderFactory = DbProviderFactories.GetFactory(ProviderName);
            connection = dbProviderFactory.CreateConnection();
            connection.ConnectionString = ConnectionString;

            return connection;

        }

        public static DbCommand CreateCommand(string QueryText, DbConnection connection)
        {
            DbCommand command = null;
            if (string.IsNullOrEmpty(QueryText) || connection == null) return null;

            command = dbProviderFactory.CreateCommand();
            command.CommandText = QueryText;
            command.Connection = connection;

            return command;

        }

        public static DbDataAdapter CreateDataAdapter(DbCommand command)
        {
            DbDataAdapter dataAdapter = null;
            if (command == null) return null;

            dataAdapter = dbProviderFactory.CreateDataAdapter();
            dataAdapter.SelectCommand = command;

            return dataAdapter;

        }

        public enum ProviderTypes
        {
            Odbc,
            OleDb,
            OracleClient,
            SqlClient
        }
    }
}
