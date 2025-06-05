 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace DatabaseOperation
{
 
    public class DBServerConnection : Microsoft.Data.ConnectionUI.DataConnectionDialog
    {
        public DBServerConnection()
        {

        }
     
        public DialogResult ShowConnectionDialog(bool IsSQL=true)
        {
            Microsoft.Data.ConnectionUI.DataSource.AddStandardDataSources(this);

            if (IsSQL)
            {
                SelectedDataSource = Microsoft.Data.ConnectionUI.DataSource.SqlDataSource;
                SelectedDataProvider = Microsoft.Data.ConnectionUI.DataProvider.SqlDataProvider;
                ConnectionString = ""; 
            }    
            
            return Microsoft.Data.ConnectionUI.DataConnectionDialog.Show(this);
        }
 
    }
}
