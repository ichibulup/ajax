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
using OfficeOpenXml;
using Xceed.Words.NET;

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
            var datatablesDB = empDB.ListAll();
            using (ExcelPackage Ep = new ExcelPackage()) {
                ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Report");
                Sheet.Cells["A1"].Value = "EmployeeID"; // [1, 1]
                Sheet.Cells["B1"].Value = "Name"; // [1, 2]
                Sheet.Cells["C1"].Value = "Age"; // [1, 3]
                Sheet.Cells["D1"].Value = "State"; // [1, 4]
                Sheet.Cells["E1"].Value = "Country"; // [1, 5]
                int row = 2;
                foreach (var item in datatablesDB) {
                    Sheet.Cells[string.Format("A{0}", row)].Value = item.EmployeeID;
                    Sheet.Cells[string.Format("B{0}", row)].Value = item.Name;
                    Sheet.Cells[string.Format("C{0}", row)].Value = item.Age;
                    Sheet.Cells[string.Format("D{0}", row)].Value = item.State;
                    Sheet.Cells[string.Format("E{0}", row)].Value = item.Country;
                    row++;
                }
                
                Sheet.Column(1).Style.Numberformat.Format = "0"; // EmployeeID
                Sheet.Column(3).Style.Numberformat.Format = "0"; // Age
                Sheet.Cells["A:AZ"].AutoFitColumns();

                // var stream = new MemoryStream();
                // excel.SaveAs(stream);
                // var fileName = "Employees.xlsx";

                // return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment: filename=" + "Report.xlsx");
                Response.BinaryWrite(Ep.GetAsByteArray());
                Response.End();
            }
        }
        
        public void ExportToWord() {
            List<Employee> employee = new List<Employee>();
            var datatablesDB = empDB.ListAll();
            using (DocX document = DocX.Create("PeopleExport.docx"))
            {
                document.InsertParagraph("Employee List").Font("Arial").FontSize(20).Bold().SpacingAfter(20);
                var table = document.InsertTable(empDB.ListAll().Count + 1, 4);

                table.Rows[0].Cells[0].Paragraphs.First().Append("Name").Bold();
                table.Rows[0].Cells[1].Paragraphs.First().Append("Age").Bold();
                table.Rows[0].Cells[2].Paragraphs.First().Append("State").Bold();
                table.Rows[0].Cells[3].Paragraphs.First().Append("country").Bold();

                int row = 1;
                foreach (var item in datatablesDB) {
                    // table.Rows[row].Cells[0].Paragraphs.First().Append(item.EmployeeID.ToString());
                    table.Rows[row].Cells[0].Paragraphs.First().Append(item.Name);
                    table.Rows[row].Cells[1].Paragraphs.First().Append(item.Age.ToString());
                    table.Rows[row].Cells[2].Paragraphs.First().Append(item.State);
                    table.Rows[row].Cells[3].Paragraphs.First().Append(item.Country);
                    row++;
                }

                // document.Save();
                //
                // Response.Clear();
                // Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                // Response.AddHeader("content-disposition", "attachment: filename=" + "Report.docx");
                // // Response.BinaryWrite(Ep.GetAsByteArray());
                // // Response.BinaryWrite(document.GetHashCode());
                // Response.BinaryWrite(document.InsertTable(table));
                // Response.End();
                
                using (MemoryStream stream = new MemoryStream()) {
                    document.SaveAs(stream);
                    stream.Position = 0;

                    Response.Clear();
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    Response.AddHeader("content-disposition", "attachment; filename=PeopleExport.docx");
                    Response.BinaryWrite(stream.ToArray());
                    Response.End();
                }
            }

            // byte[] fileBytes = System.IO.File.ReadAllBytes("PeopleExport.docx");
            // return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "PeopleExport.docx");
        }
    }
}