using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Manipulation2
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = GetConnectionString();
            ReturnIdentity(connectionString);
        }
        private static void ReturnIdentity(string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                
                SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM dbo.Excel1", connection);

                // Create a SqlCommand to execute the stored procedure.
                adapter.InsertCommand = new SqlCommand("GetExcel", connection);
                adapter.InsertCommand.CommandType = CommandType.StoredProcedure;

                              

               // Create a DataTable and fill it.
                DataTable categories = new DataTable();
                adapter.Fill(categories);

                

                // Update the database.
                adapter.Update(categories);

                foreach (DataRow row in categories.Rows)
                {
                    Console.WriteLine("  {0}: {1}", row[0], row[1]);
                }
            }
        }

        static private string GetConnectionString()
        {
            return "Data Source=(local);Initial Catalog=Northwind;Integrated Security=true";
        }
    }
}
