using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GE_Merchant_Picker
{
    class DAL
    {
        static public String readFromSQL(String query, String columnName, string connectionString)
        {
            string SQLResult = string.Empty;

            using (SqlConnection connection = new SqlConnection())
            {
                connection.ConnectionString = connectionString;
                connection.Open();
                using (SqlCommand myCommand = new SqlCommand(query, connection))
                using (SqlDataReader myReader = myCommand.ExecuteReader())
                {
                    while (myReader.Read())
                    {
                        SQLResult = myReader[columnName].ToString();
                    }
                }
            }

            return SQLResult;

        }
    }
}
