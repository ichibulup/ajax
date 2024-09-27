using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace CRUDAjax.Models
{
    // public class EmployeeDB
    // {
    //     //declare connection string
    //     string cs = ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString;

    //     //Return list of all Employees
    //     public List<Employee> ListAll()
    //     {
    //         List<Employee> lst = new List<Employee>();
    //         using(SqlConnection con=new SqlConnection(cs))
    //         {
    //             con.Open();
    //             SqlCommand com = new SqlCommand("SelectEmployee",con);
    //             com.CommandType = CommandType.StoredProcedure;
    //             SqlDataReader rdr = com.ExecuteReader();
    //             while(rdr.Read())
    //             {
    //                 lst.Add(new Employee { 
    //                     EmployeeID=Convert.ToInt32(rdr["EmployeeId"]),
    //                     Name=rdr["Name"].ToString(),
    //                     Age = Convert.ToInt32(rdr["Age"]),
    //                     State = rdr["State"].ToString(),
    //                     Country = rdr["Country"].ToString(),
    //                 });
    //             }
    //             return lst;
    //         }
    //     }

    //     //Method for Adding an Employee
    //     public int Add(Employee emp)
    //     {
    //         int i;
    //         using(SqlConnection con=new SqlConnection(cs))
    //         {
    //             con.Open();
    //             SqlCommand com = new SqlCommand("InsertUpdateEmployee", con);
    //             com.CommandType = CommandType.StoredProcedure;
    //             com.Parameters.AddWithValue("@Id",emp.EmployeeID);
    //             com.Parameters.AddWithValue("@Name", emp.Name);
    //             com.Parameters.AddWithValue("@Age", emp.Age);
    //             com.Parameters.AddWithValue("@State", emp.State);
    //             com.Parameters.AddWithValue("@Country", emp.Country);
    //             com.Parameters.AddWithValue("@Action", "Insert");
    //             i = com.ExecuteNonQuery();
    //         }
    //         return i;
    //     }

    //     //Method for Updating Employee record
    //     public int Update(Employee emp)
    //     {
    //         int i;
    //         using (SqlConnection con = new SqlConnection(cs))
    //         {
    //             con.Open();
    //             SqlCommand com = new SqlCommand("InsertUpdateEmployee", con);
    //             com.CommandType = CommandType.StoredProcedure;
    //             com.Parameters.AddWithValue("@Id", emp.EmployeeID);
    //             com.Parameters.AddWithValue("@Name", emp.Name);
    //             com.Parameters.AddWithValue("@Age", emp.Age);
    //             com.Parameters.AddWithValue("@State", emp.State);
    //             com.Parameters.AddWithValue("@Country", emp.Country);
    //             com.Parameters.AddWithValue("@Action", "Update");
    //             i = com.ExecuteNonQuery();
    //         }
    //         return i;
    //     }

    //     //Method for Deleting an Employee
    //     public int Delete(int ID)
    //     {
    //         int i;
    //         using (SqlConnection con = new SqlConnection(cs))
    //         {
    //             con.Open();
    //             SqlCommand com = new SqlCommand("DeleteEmployee", con);
    //             com.CommandType = CommandType.StoredProcedure;
    //             com.Parameters.AddWithValue("@Id", ID);
    //             i = com.ExecuteNonQuery();
    //         }
    //         return i;
    //     }
    // }

    public class EmployeeDB
    {
        private readonly string connectionString = "Data Source=ADMIN\\SQLEXPRESS;Initial Catalog=MVCDB;Integrated Security=True;Trusted_Connection=True";

        public List<Employee> ListAll()
        {
            List<Employee> employees = new List<Employee>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = "SELECT * FROM Employee";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    employees.Add(new Employee
                    {
                        EmployeeID = (int)reader["EmployeeID"],
                        Name = reader["Name"].ToString(),
                        Age = (int)reader["Age"],
                        State = reader["State"].ToString(),
                        Country = reader["Country"].ToString()
                    });
                }
            }

            return employees;
        }

        public bool Add(Employee employee)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = "INSERT INTO Employee (Name, Age, State, Country) VALUES (@Name, @Age, @State, @Country)";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@Name", employee.Name);
                cmd.Parameters.AddWithValue("@Age", employee.Age);
                cmd.Parameters.AddWithValue("@State", employee.State);
                cmd.Parameters.AddWithValue("@Country", employee.Country);
                conn.Open();
                return cmd.ExecuteNonQuery() > 0;
            }
        }

        public bool Update(Employee employee)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = "UPDATE Employee SET Name = @Name, Age = @Age, State = @State, Country = @Country WHERE EmployeeID = @EmployeeID";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@EmployeeID", employee.EmployeeID);
                cmd.Parameters.AddWithValue("@Name", employee.Name);
                cmd.Parameters.AddWithValue("@Age", employee.Age);
                cmd.Parameters.AddWithValue("@State", employee.State);
                cmd.Parameters.AddWithValue("@Country", employee.Country);
                conn.Open();
                return cmd.ExecuteNonQuery() > 0;
            }
        }

        public bool Delete(int id)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = "DELETE FROM Employee WHERE EmployeeID = @EmployeeID";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@EmployeeID", id);
                conn.Open();
                return cmd.ExecuteNonQuery() > 0;
            }
        }
    }
}
