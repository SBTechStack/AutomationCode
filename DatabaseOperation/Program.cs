using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseOperation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            DBServerConnection dBServerConnection = new DBServerConnection();
            dBServerConnection.ShowConnectionDialog(true);
            DatabaseActivity.DatabaseActivitys(dBServerConnection.CompanyName, dBServerConnection.ConnectionString, "SqlClient", "select *from tblPRLEmployee");
        }
    }
}
