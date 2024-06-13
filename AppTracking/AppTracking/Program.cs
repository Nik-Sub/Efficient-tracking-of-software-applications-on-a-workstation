using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AppTracking.domain;
using AppTracking.forms;
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
            //AllocConsole();

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

        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

    }

}
