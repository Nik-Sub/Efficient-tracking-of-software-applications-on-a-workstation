using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppTracking.domain
{
    internal class ExcelPrinterImpl : PrinterImpl
    {
        public void printReport(List<Dictionary<string, string>> apps)
        {
            // Create a new Excel application
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // Add a workbook
            Workbook workbook = excelApp.Workbooks.Add();
            Worksheet worksheet = workbook.Worksheets[1];

            // Set column headers
            worksheet.Cells[1, 1] = "Display Name";
            worksheet.Cells[1, 2] = "Install Date";
            worksheet.Cells[1, 3] = "Version";
            worksheet.Cells[1, 4] = "UpdateID";
            worksheet.Cells[1, 5] = "UpdateDescription";
            worksheet.Cells[1, 6] = "UpdateInstallDate";
            //worksheet.Cells[1, 4] = "Install Location";

            // Write data to the Excel sheet
            int row = 2;
            foreach (var app in apps)
            {
                worksheet.Cells[row, 1] = app["DisplayName"];
                string formattedDate = "";
                if (app["InstallDate"] != null)
                {
                    try
                    {
                        DateTime date = DateTime.ParseExact(app["InstallDate"], "yyyyMMdd", CultureInfo.InvariantCulture);
                        formattedDate = date.ToString("dd-MM-yyyy");
                    }
                    catch (Exception e)
                    {
                        formattedDate = "";
                    }
                }
                worksheet.Cells[row, 2] = formattedDate;
                worksheet.Cells[row, 3] = app["DisplayVersion"];
                worksheet.Cells[row, 4] = app["UpdateID"];
                worksheet.Cells[row, 5] = app["UpdateDescription"];
                worksheet.Cells[row, 6] = app["UpdateInstallDate"];
                //worksheet.Cells[row, 4] = app["InstallLocation"];
                row++;
            }

            // Save the workbook and close Excel
            string fileName = "output1.xlsx";
            //Console.WriteLine(System.AppDomain.CurrentDomain.BaseDirectory);
            string filePath = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, fileName);
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            workbook.SaveAs(filePath);
            workbook.Close();
            excelApp.Quit();

            // Release COM objects to avoid memory leaks
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("Data exported to Excel successfully.");
        }
    }
}
