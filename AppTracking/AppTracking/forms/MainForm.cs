using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppTracking.forms
{
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
            DataTable dataTable = new DataTable();
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
