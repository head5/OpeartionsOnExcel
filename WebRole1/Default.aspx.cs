using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace WebRole1
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void butShowDetails_Click(object sender, EventArgs e)
        {
            Workbook workBook;
            SharedStringTable sharedStrings;
            IEnumerable<Sheet> workSheets;
            WorksheetPart EmpSheet;
            List<Employee> Employees;
            string empID;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(@"C:\Users\v-prwagh\Desktop\Employee.xlsx", true))
            {
                workBook = document.WorkbookPart.Workbook;
                workSheets = workBook.Descendants<Sheet>();
                sharedStrings = document.WorkbookPart.SharedStringTablePart.SharedStringTable;

                empID = workSheets.First(s => s.Name == @"EmployeeDetails").Id;
                EmpSheet = (WorksheetPart)document.WorkbookPart.GetPartById(empID);
                Employees = Employee.LoadEmployee(EmpSheet.Worksheet, sharedStrings);

                IEnumerable<Employee> allemployees = from employee in Employees select employee;
                string result = null;
                result = result + "All Employees \n";
                result = result + "Emp_ID" + "          Emp_Name" + "           Compentency" + "            Location\n";



                //foreach (Employee c in allemployees)
                //{
                //    result = result + c.Emp_ID + "          " + c.Emp_Name + "          " + c.Compentency + "            " + c.Location + "\n";
                //}

                var selectedEmployee = from employee in Employees
                                       where employee.Compentency == "C2"
                                       select new { employee.Emp_ID, employee.Emp_Name, employee.Compentency, employee.Location };


                foreach (var c in selectedEmployee)
                {
                    result = result + c.Emp_ID + "          " + c.Emp_Name + "          " + c.Compentency + "            " + c.Location + "\n";
                }

              // TextResult.Text = result;


                tblResult.Controls.Clear();
                tblResult.BorderStyle = BorderStyle.Inset;
                tblResult.BorderWidth = Unit.Pixel(2);
                foreach (Employee c in allemployees)
                {
                    TableRow rowNew = new TableRow();
                    
                    rowNew.BorderColor = System.Drawing.Color.Blue;
                    tblResult.Controls.Add(rowNew);

                    TableCell cellNewEmpID = new TableCell();
                    TableCell cellNewName = new TableCell();
                    TableCell cellNewCompentency = new TableCell();
                    TableCell cellNewLocation = new TableCell();
                    cellNewEmpID.Text = c.Emp_ID;
                    cellNewName.Text = c.Emp_Name;
                    cellNewCompentency.Text = c.Compentency;
                    cellNewLocation.Text = c.Location;
                    
                    //cellNewEmpID.BorderWidth = 1;
                    //cellNewName.BorderWidth = 1;
                    //cellNewCompentency.BorderWidth = 1;
                    //cellNewLocation.BorderWidth = 1;

                    rowNew.Controls.Add(cellNewEmpID);
                    rowNew.Controls.Add(cellNewName);
                    rowNew.Controls.Add(cellNewCompentency);
                    rowNew.Controls.Add(cellNewLocation);
                        //result = result + c.Emp_ID + "          " + c.Emp_Name + "          " + c.Compentency + "            " + c.Location + "\n";                 

                }

                Outlook.Application oApp = new Outlook.Application();
                Outlook._MailItem oMailItem = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                oMailItem.To = "v-prwagh@microsoft.com";
                oMailItem.CC = "";
                oMailItem.Body = "Trial Mail Body .................";
                oMailItem.Subject = "Trial Mail";
                oMailItem.Display(true);
            }
        }
    }
}