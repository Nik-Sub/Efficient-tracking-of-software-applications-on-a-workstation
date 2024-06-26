using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
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
        static void Main(string[] args)
        {
            // Allocate a console for the application
            //
            /*// Get the current culture and UI culture
            CultureInfo currentCulture = CultureInfo.CurrentCulture;
            CultureInfo currentUICulture = CultureInfo.CurrentUICulture;

            // Print the culture information
            Console.WriteLine("Current Culture: " + currentCulture.Name);
            Console.WriteLine("Current UI Culture: " + currentUICulture.Name);
            Console.WriteLine("Date format (short): " + currentCulture.DateTimeFormat.ShortDatePattern);
            Console.WriteLine("Date format (long): " + currentCulture.DateTimeFormat.LongDatePattern);
            Console.WriteLine("Time format: " + currentCulture.DateTimeFormat.LongTimePattern);*/

            /*// Instantiate your form
            Form1 form = new Form1();

            // Run the form
            System.Windows.Forms.Application.Run(form);*/
            //AllocConsole();
            //Console.WriteLine(args.Length);
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new MainForm(args));
            


            /*ApplicationConfiguration.Initialize();
            Application.Run(new Form1());*/

            /*AppReader appReader = new AppReader(new WinAppReaderImplementation());
            List<Dictionary<string, string>> apps = appReader.getAppl();

            Console.WriteLine(apps.Count);*/

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
