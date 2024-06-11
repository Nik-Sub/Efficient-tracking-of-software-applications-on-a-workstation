using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace AppTracking
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Allocate a console for the application
            AllocConsole();

            /*// Instantiate your form
            Form1 form = new Form1();

            // Run the form
            System.Windows.Forms.Application.Run(form);*/

            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new MainForm());

            /*ApplicationConfiguration.Initialize();
            Application.Run(new Form1());*/

            AppReader appReader = new AppReader(new WinAppReaderImplementation());
            List<Dictionary<string, string>> apps = appReader.getAppl();

            Console.WriteLine(apps.Count);

            // Print the installed applications
            /*foreach (var app in apps)
            {
                Console.WriteLine($"Display Name: {app["DisplayName"]}");
                Console.WriteLine($"Install Date: {app["InstallDate"]}");
                Console.WriteLine($"Version: {app["DisplayVersion"]}");
                Console.WriteLine($"InstallLocation: {app["InstallLocation"]}");
                Console.WriteLine();
            }*/

            // Create a new Excel application
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // Add a workbook
            Workbook workbook = excelApp.Workbooks.Add();
            Worksheet worksheet = workbook.Worksheets[1];

            // Set column headers
            worksheet.Cells[1, 1] = "Display Name";
            worksheet.Cells[1, 2] = "Install Date";
            worksheet.Cells[1, 3] = "Version";
            worksheet.Cells[1, 4] = "Install Location";

            // Write data to the Excel sheet
            int row = 2;
            foreach (var app in apps)
            {
                worksheet.Cells[row, 1] = app["DisplayName"];
                worksheet.Cells[row, 2] = app["InstallDate"];
                worksheet.Cells[row, 3] = app["DisplayVersion"];
                worksheet.Cells[row, 4] = app["InstallLocation"];
                row++;
            }

            // Save the workbook and close Excel
            string fileName = "output.xlsx";
            //Console.WriteLine(System.AppDomain.CurrentDomain.BaseDirectory);
            string filePath = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, fileName);
            workbook.SaveAs(filePath);
            workbook.Close();
            excelApp.Quit();

            // Release COM objects to avoid memory leaks
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("Data exported to Excel successfully.");

            Console.ReadLine();
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();


        public class MainForm : Form
        {
            private DataGridView dataGridView;

            public MainForm()
            {
                this.Text = "Installed Applications";
                this.Size = new Size(800, 600);

                dataGridView = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells // This line auto sizes the columns
                };
                this.Controls.Add(dataGridView);

                AppReader appReader = new AppReader(new WinAppReaderImplementation());
                List<Dictionary<string, string>> apps = appReader.getAppl();

                // Create a DataTable to hold the data
                System.Data.DataTable dataTable = new System.Data.DataTable();
                dataTable.Columns.Add("Display Name");
                dataTable.Columns.Add("Install Date");
                dataTable.Columns.Add("Version");
                dataTable.Columns.Add("Install Location");

                // Populate the DataTable with your data
                foreach (var app in apps)
                {
                    DataRow row = dataTable.NewRow();
                    row["Display Name"] = app["DisplayName"];
                    row["Install Date"] = app["InstallDate"];
                    row["Version"] = app["DisplayVersion"];
                    row["Install Location"] = app["InstallLocation"];
                    dataTable.Rows.Add(row);
                }

                // Set the DataSource of the DataGridView to the DataTable
                dataGridView.DataSource = dataTable;
            }
        }
    }

}
