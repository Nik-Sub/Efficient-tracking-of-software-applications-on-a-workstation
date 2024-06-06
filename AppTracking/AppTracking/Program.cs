using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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

            /*ApplicationConfiguration.Initialize();
            Application.Run(new Form1());*/

            AppReader appReader = new AppReader(new WinAppReaderImplementation());
            List<Dictionary<string, string>> apps = appReader.getAppl();

            Console.WriteLine(apps.Count);

            // Print the installed applications
            foreach (var app in apps)
            {
                Console.WriteLine($"Display Name: {app["DisplayName"]}");
                Console.WriteLine($"Install Date: {app["InstallDate"]}");
                Console.WriteLine($"Version: {app["DisplayVersion"]}");
                Console.WriteLine($"InstallLocation: {app["InstallLocation"]}");
                Console.WriteLine();
            }

            Console.ReadLine();
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
    }
}
