using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace ConsoleTestApp
{
    public class ExcelHelper
    {

        public void GenerateExcel()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };


            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            // Select the Excel cells, in the range c1 to c7 in the worksheet.
            Range aRange = ws.get_Range("C1", "C7");

            if (aRange == null)
            {
                Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            }

            // Fill the cells in the C1 to C7 range of the worksheet with the number 6.
            Object[] args = new Object[1];
            args[0] = 6;
            aRange.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, aRange, args);

            // Change the cells in the C1 to C7 range of the worksheet to the number 8.
            aRange.Value2 = 8;
        }

        public void CreateExcelFile(string fileName)
        {
            var customers = GetCustomerData(10);
            string[] columns = { "ID", "Name", "Address", "Phone"};

            var excel = new Microsoft.Office.Interop.Excel.Application() { Visible = false, DisplayAlerts = false };

            var worKbooK = excel.Workbooks.Add(Type.Missing);
            var worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
            worKsheeT.Name = "CustomerReport";

            //Merge the cells from [1,1] to [1,8] depending on requirements as in the following:
            worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, columns.Length]].Merge();

            worKsheeT.Cells[1, 1] = "Customer Report Card";
            worKsheeT.Cells.Font.Size = 15;

            int rowcount = 2;            
            Range celLrangE;

            //For Columns Names 
            for (int i = 1; i <= columns.Length; i++)
            {
                worKsheeT.Cells[2, i] = columns[i - 1];
                worKsheeT.Cells.Font.Color = System.Drawing.Color.Black;
            }

            //For Data 
            foreach (var datarow in customers)
            {
                rowcount += 1;
                for (int i = 1; i <= columns.Length; i++)
                {
                    worKsheeT.Cells[rowcount, i] = GetCustomerValue(datarow, i);
                }
            }

            celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, columns.Length]];
            celLrangE.EntireColumn.AutoFit();
            Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, columns.Length]];

            worKbooK.SaveAs(fileName); ;
            worKbooK.Close();
            excel.Quit();

        }

        public IList<Customer> GetCustomerData(int noOfRows)
        {
            List<Customer> customers = new List<Customer>();

            for(int i= 0; i <= noOfRows; i++)
            {
                var cus = new Customer() { ID = 100+i, Name = "Test "+i, Address = "BNG", Phone = 7238947 };
                customers.Add(cus);
            }

            return customers;
        }

        public string GetCustomerValue(Customer customer, int position)
        {
            switch (position)
            {
                case 1: return customer.ID.ToString();
                case 2: return customer.Name.ToString();
                case 3: return customer.Address.ToString();
                case 4: return customer.Phone.ToString();
                default: return customer.Phone.ToString();
            }
        }
    }

    public class Customer
    {
        public int ID { get; set; }
        public string Name { get; set; }

        public string Address { get; set; }

        public int Phone { get; set; }
    }
}
