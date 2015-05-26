using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebRole1
{
    public class Employee
    {
        private string _emp_ID;
        private string _emp_Name;
        private string _compentency;
        private string _location;

        public string Emp_ID
        {
            get { return _emp_ID; }
            set { _emp_ID = value; }
        }
        public string Emp_Name
        {
            get { return _emp_Name; }
            set { _emp_Name = value; }
        }
        public string Compentency
        {
            get { return _compentency; }
            set { _compentency = value; }
        }
        public string Location
        {
            get { return _location; }
            set { _location = value; }
        }



        public static List<Employee> LoadEmployee(Worksheet worksheet, SharedStringTable sharedString)
        {
            //Initialize the customer list.
            List<Employee> result = new List<Employee>();

            //LINQ query to skip first row with column names.
            IEnumerable<Row> dataRows = from row in worksheet.Descendants<Row>() where row.RowIndex > 1 select row;

            foreach (Row row in dataRows)
            {
                //LINQ query to return the row's cell values.
                //Where clause filters out any cells that do not contain a value.
                //Select returns the value of a cell unless the cell contains
                //  a Shared String.
                //If the cell contains a Shared String, its value will be a 
                //  reference id which will be used to look up the value in the 
                //  Shared String table.
                IEnumerable<String> textValues =
                  from cell in row.Descendants<Cell>()
                  where cell.CellValue != null
                  select
                    (cell.DataType != null
                      && cell.DataType.HasValue
                      && cell.DataType == CellValues.SharedString
                    ? sharedString.ChildElements[
                      int.Parse(cell.CellValue.InnerText)].InnerText
                    : cell.CellValue.InnerText)
                  ;

                //Check to verify the row contained data.
                if (textValues.Count() > 0)
                {
                    //Create a customer and add it to the list.
                    var textArray = textValues.ToArray();
                    Employee employee = new Employee();
                    employee.Emp_ID = textArray[0];
                    employee.Emp_Name = textArray[1];
                    employee.Compentency = textArray[2];
                    employee.Location = textArray[3];
                    result.Add(employee);
                }
                else
                {
                    //If no cells, then you have reached the end of the table.
                    break;
                }
            }

            //Return populated list of customers.
            return result;
        }
    }
}