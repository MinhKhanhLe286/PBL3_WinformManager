using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KhachSan
{
    internal class DataProvider
    {
        private String connectString = "Data Source=DESKTOP-ENHGL7D\\SQLEXPRESS;Initial Catalog=myHotel1;Integrated Security=True";


        public DataTable execQuery(string query)
        {
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(connectString))
            {
                con.Open();

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = query;

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dt);
                con.Close();

            }
            return dt;
        }

        public int execNonQuery(String query)
        {
            int data = 0;
            try
            {

                using (SqlConnection conn = new SqlConnection(connectString))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(query, conn);
                    data = cmd.ExecuteNonQuery();
                    conn.Close();
                }
                return data;
            }
            catch (Exception ex)
            {
                // Log the exception or handle it as needed
                Console.WriteLine(ex.Message);
                return data;
            }
        }

        public object execScaler(String query)
        {
            object data = 0;
            using (SqlConnection conn = new SqlConnection(connectString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(query, conn);
                data = cmd.ExecuteScalar();
                conn.Close();
            }
            return data;
        }
    }
}
