using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Data.SqlClient;
using CRUDAjax.Models;
using System.IO;
using OfficeOpenXml;
using System.Windows.Forms;
using System.Text.Json;
using System.Configuration;
using System.Reflection;
using System.Data;
using System.Text;

namespace CRUDAjax.Controllers
{
    // public class HomeController : Controller
    // {
    //     EmployeeDB empDB = new EmployeeDB();
    //     // GET: Home
    //     public ActionResult Index()
    //     {
    //         return View();
    //     }
    //     public JsonResult List()
    //     {
    //         return Json(empDB.ListAll(),JsonRequestBehavior.AllowGet);
    //     }
    //     public JsonResult Add(Employee emp)
    //     {
    //         return Json(empDB.Add(emp), JsonRequestBehavior.AllowGet);
    //     }
    //     public JsonResult GetbyID(int ID)
    //     {
    //         var Employee = empDB.ListAll().Find(x => x.EmployeeID.Equals(ID));
    //         return Json(Employee, JsonRequestBehavior.AllowGet);
    //     }
    //     public JsonResult Update(Employee emp)
    //     {
    //         return Json(empDB.Update(emp), JsonRequestBehavior.AllowGet);
    //     }
    //     public JsonResult Delete(int ID)
    //     {
    //         return Json(empDB.Delete(ID), JsonRequestBehavior.AllowGet);
    //     }
    // }

    // public class HomeController : Controller
    // {
    //     private readonly string connectionString = "your_connection_string_here";

    //     // Function to Load Data
    //     public ActionResult Load()
    //     {
    //         List<Employee> employees = new List<Employee>();

    //         using (SqlConnection conn = new SqlConnection(connectionString))
    //         {
    //             string query = "SELECT * FROM Employee";
    //             SqlCommand cmd = new SqlCommand(query, conn);
    //             conn.Open();
    //             SqlDataReader reader = cmd.ExecuteReader();

    //             while (reader.Read())
    //             {
    //                 employees.Add(new Employee
    //                 {
    //                     EmployeeID = (int)reader["EmployeeID"],
    //                     Name = reader["Name"].ToString(),
    //                     Age = (int)reader["Age"],
    //                     State = reader["State"].ToString(),
    //                     Country = reader["Country"].ToString()
    //                 });
    //             }
    //         }

    //         return Json(new { data = employees }, JsonRequestBehavior.AllowGet);
    //     }

    //     // Add Employee
    //     [HttpPost]
    //     public ActionResult Add(Employee employee)
    //     {
    //         using (SqlConnection conn = new SqlConnection(connectionString))
    //         {
    //             string query = "INSERT INTO Employee (Name, Age, State, Country) VALUES (@Name, @Age, @State, @Country)";
    //             SqlCommand cmd = new SqlCommand(query, conn);
    //             cmd.Parameters.AddWithValue("@Name", employee.Name);
    //             cmd.Parameters.AddWithValue("@Age", employee.Age);
    //             cmd.Parameters.AddWithValue("@State", employee.State);
    //             cmd.Parameters.AddWithValue("@Country", employee.Country);
    //             conn.Open();
    //             cmd.ExecuteNonQuery();
    //         }
    //         return Json(new { success = true });
    //     }

    //     // Update Employee
    //     [HttpPost]
    //     public ActionResult Update(Employee employee)
    //     {
    //         using (SqlConnection conn = new SqlConnection(connectionString))
    //         {
    //             string query = "UPDATE Employee SET Name = @Name, Age = @Age, State = @State, Country = @Country WHERE EmployeeID = @EmployeeID";
    //             SqlCommand cmd = new SqlCommand(query, conn);
    //             cmd.Parameters.AddWithValue("@EmployeeID", employee.EmployeeID);
    //             cmd.Parameters.AddWithValue("@Name", employee.Name);
    //             cmd.Parameters.AddWithValue("@Age", employee.Age);
    //             cmd.Parameters.AddWithValue("@State", employee.State);
    //             cmd.Parameters.AddWithValue("@Country", employee.Country);
    //             conn.Open();
    //             cmd.ExecuteNonQuery();
    //         }
    //         return Json(new { success = true });
    //     }

    //     // Delete Employee
    //     [HttpPost]
    //     public ActionResult Delete(int id)
    //     {
    //         using (SqlConnection conn = new SqlConnection(connectionString))
    //         {
    //             string query = "DELETE FROM Employee WHERE EmployeeID = @EmployeeID";
    //             SqlCommand cmd = new SqlCommand(query, conn);
    //             cmd.Parameters.AddWithValue("@EmployeeID", id);
    //             conn.Open();
    //             cmd.ExecuteNonQuery();
    //         }
    //         return Json(new { success = true });
    //     }

    //     // Search Employee
    //     [HttpGet]
    //     public ActionResult Search(string searchTerm)
    //     {
    //         List<Employee> employees = new List<Employee>();

    //         using (SqlConnection conn = new SqlConnection(connectionString))
    //         {
    //             string query = "SELECT * FROM Employee WHERE Name LIKE '%' + @searchTerm + '%' OR Country LIKE '%' + @searchTerm + '%'";
    //             SqlCommand cmd = new SqlCommand(query, conn);
    //             cmd.Parameters.AddWithValue("@searchTerm", searchTerm);
    //             conn.Open();
    //             SqlDataReader reader = cmd.ExecuteReader();

    //             while (reader.Read())
    //             {
    //                 employees.Add(new Employee
    //                 {
    //                     EmployeeID = (int)reader["EmployeeID"],
    //                     Name = reader["Name"].ToString(),
    //                     Age = (int)reader["Age"],
    //                     State = reader["State"].ToString(),
    //                     Country = reader["Country"].ToString()
    //                 });
    //             }
    //         }

    //         return Json(new { data = employees }, JsonRequestBehavior.AllowGet);
    //     }

    //     // Export Data to Excel (Placeholder)
    //     public ActionResult ExportToExcel()
    //     {
    //         // TODO: Add logic to export to Excel
    //         return Json(new { success = true });
    //     }

    //     // Export Data to Word (Placeholder)
    //     public ActionResult ExportToWord()
    //     {
    //         // TODO: Add logic to export to Word
    //         return Json(new { success = true });
    //     }
    // }
    public class HomeController : Controller
    {
        EmployeeDB empDB = new EmployeeDB();

        public ActionResult Index()
        {
            return View();
        }

        public JsonResult List()
        {
            return Json(empDB.ListAll(), JsonRequestBehavior.AllowGet);
            //try
            //{
            //    return Json(empDB.ListAll(), JsonRequestBehavior.AllowGet);
            //}
            //catch (Exception ex)
            //{
            //    return Json(new { success = false, message = ex.Message }, JsonRequestBehavior.AllowGet);
            //}
        }

        public JsonResult Add(Employee emp)
        {
            return Json(empDB.Add(emp), JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetbyID(int ID)
        {
            var employee = empDB.ListAll().Find(x => x.EmployeeID.Equals(ID));
            return Json(employee, JsonRequestBehavior.AllowGet);
        }

        public JsonResult Update(Employee emp)
        {
            return Json(empDB.Update(emp), JsonRequestBehavior.AllowGet);
        }

        public JsonResult Delete(int ID)
        {
            return Json(empDB.Delete(ID), JsonRequestBehavior.AllowGet);
        }

        //private void frmExportData_Load(object sender, EventArgs e) {
        //    // TODO: This line of code loads data into the 'appData.Customers' table. You can move, or remove it, as needed.
        //    this.customersTableAdapter.Fill(this.appData.Customers);
        //    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //}

        public void ExportToExcel() {
            //using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" }) {
            //    if (sfd.ShowDialog() == DialogResult.OK) {
            //        var fileInfo = new FileInfo(sfd.FileName);
            //        using (ExcelPackage excelPackage = new ExcelPackage(fileInfo)) {
            //            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Customers");
            //            // Load the data from the DataTable to the Excel worksheet
            //            worksheet.Cells.LoadFromDataTable(this.appData.Customers.CopyToDataTable(), true);
            //            // Save the Excel file
            //            excelPackage.Save();
            //        }
            //    }
            //}
            var datatablesDB = empDB.ListAll();
            using (ExcelPackage Ep = new ExcelPackage()) {
                ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Report");
                Sheet.Cells["A1"].Value = "Name";
                Sheet.Cells["B1"].Value = "Age";
                Sheet.Cells["C1"].Value = "State";
                Sheet.Cells["D1"].Value = "Country";
                int row = 2;
                foreach (var item in datatablesDB) {
                    Sheet.Cells[string.Format("A{0}", row)].Value = item.Name;
                    Sheet.Cells[string.Format("B{0}", row)].Value = item.Age;
                    Sheet.Cells[string.Format("C{0}", row)].Value = item.State;
                    Sheet.Cells[string.Format("D{0}", row)].Value = item.Country;
                    row++;
                }
                Sheet.Cells["A:AZ"].AutoFitColumns();
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment: filename=" + "Report.xlsx");
                Response.BinaryWrite(Ep.GetAsByteArray());
                Response.End();
            }
        }

        public void ExportToWord() {
            Export();
        }

        public FileResult Export() {
            var customers = empDB.ListAll();

            //Building an HTML string.
            StringBuilder sb = new StringBuilder();

            //Table start.
            sb.Append("<table border='1' cellpadding='5' cellspacing='0' style='border: 1px solid #ccc;font-family: Arial;'>");

            //Building the Header row.
            sb.Append("<tr>");
            sb.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>CustomerID</th>");
            sb.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>ContactName</th>");
            sb.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>City</th>");
            sb.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>Country</th>");
            sb.Append("</tr>");

            //Building the Data rows.
            //fo


            foreach (var item in customers)
            {
                sb.Append("<tr>");
                //foreach (var jtem in item)
                //{
                //    sb.Append("<td style='border: 1px solid #ccc'>");
                //    sb.Append(jtem);
                //    sb.Append("</td>");
                //}
                sb.Append("</tr>");
            }

            //Table end.
            sb.Append("</table>");

            return File(Encoding.UTF8.GetBytes(sb.ToString()), "application/vnd.ms-word", "Grid.doc");
        }
    }
}