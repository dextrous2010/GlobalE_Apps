using System;
using System.Data.SqlClient;

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
                //SqlCommand com = new SqlCommand();
                //com.CommandTimeout = 1;
                try
                {
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
                catch(Exception e)
                {
                }
            }

            return SQLResult;
        }
    }
}
