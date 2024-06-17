using AppTracking.domain;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace AppTracking.forms
{
    public class MainForm : Form
    {
        private DataGridView dataGridView;
        private Button button1;
        private List<Dictionary<string, string>> apps = null;
        private TextBox filterTextBox;
        private Button filterButton;

        public MainForm()
        {
            this.Text = "Installed Applications";
            this.Size = new Size(800, 600);

            // Create a TableLayoutPanel with three columns
            TableLayoutPanel tableLayoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4
            };
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60)); // 60% for the DataGridView
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20)); // 20% for the filter text box
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20)); // 20% for the filter button
            this.Controls.Add(tableLayoutPanel);

            // Create and configure the DataGridView
            dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill // Fill the available space
            };
            tableLayoutPanel.Controls.Add(dataGridView, 0, 0);

            AppReader appReader = new AppReader(new WinAppReaderImplementation());
            apps = appReader.getAppl();

            // Create a DataTable to hold the data
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Display Name");
            dataTable.Columns.Add("Install Date");
            dataTable.Columns.Add("Version");
            dataTable.Columns.Add("UpdateID");
            dataTable.Columns.Add("UpdateDescription");
            dataTable.Columns.Add("UpdateInstallDate");

            // Populate the DataTable with your data
            foreach (var app in apps)
            {
                DataRow row = dataTable.NewRow();
                row["Display Name"] = app["DisplayName"];
                string formattedDate = "";
                if (app["InstallDate"] != null)
                {
                    try
                    {
                        DateTime date = DateTime.ParseExact(app["InstallDate"], "yyyyMMdd", CultureInfo.InvariantCulture);
                        formattedDate = date.ToString("dd-MM-yyyy");
                    }
                    catch(Exception e)
                    {
                        formattedDate = "";
                    }
                }
                row["Install Date"] = formattedDate;
                row["Version"] = app["DisplayVersion"];
                row["UpdateID"] = app["UpdateID"];
                row["UpdateDescription"] = app["UpdateDescription"];
                row["UpdateInstallDate"] = app["UpdateInstallDate"];


                dataTable.Rows.Add(row);
            }

            // Set the DataSource of the DataGridView to the DataTable
            dataGridView.DataSource = dataTable;

            // Create and configure the button
            button1 = new Button
            {
                Text = "Print",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleCenter
            };

            button1.Click += Button1_Click;

            // Create and configure the filter TextBox
            filterTextBox = new TextBox
            {
                PlaceholderText = "Enter filter text",
                Dock = DockStyle.Fill
            };

            // Create and configure the filter Button
            filterButton = new Button
            {
                Text = "Filter",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleCenter
            };
            filterButton.Click += FilterButton_Click;

            // Add controls to the TableLayoutPanel
            tableLayoutPanel.Controls.Add(filterTextBox, 1, 0);
            tableLayoutPanel.Controls.Add(filterButton, 2, 0);
            tableLayoutPanel.Controls.Add(button1, 3, 0);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Printer printer = new Printer(new ExcelPrinterImpl());
            printer.printReport(apps);

            MessageBox.Show("printReport is finished", "Print Report Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void FilterButton_Click(object sender, EventArgs e)
        {
            // Get the filter text from the TextBox
            string filterText = filterTextBox.Text.ToLower();

            // Filter the apps list based on the filter text
            List<Dictionary<string, string>> filteredApps = apps.Where(app =>
                app["DisplayName"].ToLower().Contains(filterText)
            ).ToList();

            // Create a new DataTable for the filtered data
            DataTable filteredDataTable = new DataTable();
            filteredDataTable.Columns.Add("Display Name");
            filteredDataTable.Columns.Add("Install Date");
            filteredDataTable.Columns.Add("Version");
            filteredDataTable.Columns.Add("UpdateID");
            filteredDataTable.Columns.Add("UpdateDescription");
            filteredDataTable.Columns.Add("UpdateInstallDate");

            // Populate the DataTable with the filtered data
            foreach (var app in filteredApps)
            {
                DataRow row = filteredDataTable.NewRow();
                row["Display Name"] = app["DisplayName"];
                string formattedDate = "";
                if (app["InstallDate"] != null)
                {
                    try
                    {
                        DateTime date = DateTime.ParseExact(app["InstallDate"], "yyyyMMdd", CultureInfo.InvariantCulture);
                        formattedDate = date.ToString("dd-MM-yyyy");
                    }
                    catch (Exception ee)
                    {
                        formattedDate = "";
                    }
                }
                row["Install Date"] = formattedDate;
                row["Version"] = app["DisplayVersion"];
                row["UpdateID"] = app["UpdateID"];
                row["UpdateDescription"] = app["UpdateDescription"];
                row["UpdateInstallDate"] = app["UpdateInstallDate"];

                filteredDataTable.Rows.Add(row);
            }

            // Set the DataSource of the DataGridView to the filtered DataTable
            dataGridView.DataSource = filteredDataTable;
        }
    }
}
