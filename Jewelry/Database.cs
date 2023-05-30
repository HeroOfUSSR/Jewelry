using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace Jewelry
{
    internal class Database
    {
        SqlConnection sqlConnection =
        new SqlConnection(@"Data Source=DESKTOP-2FNBGD6;Initial Catalog=Jews;Integrated Security=True");
        public void openConnection()
        {
            if (sqlConnection.State == System.Data.ConnectionState.Closed)
            {
                sqlConnection.Open();
            }

        }

        public void CloseConnection()
        {
            if (sqlConnection.State == System.Data.ConnectionState.Open)
            {
                sqlConnection.Close();
            }

        }
        public SqlConnection GetSqlConnection()
        {
            return sqlConnection;
        }
    }
}
