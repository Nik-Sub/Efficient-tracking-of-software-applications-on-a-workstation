using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Management;

class WinAppReaderImplementation : AppReaderImpl
{
    private List<string> listWithPathsCurrentUser = new List<string>();
    private List<string> listWithPathsLocalMachine = new List<string>();
    private Dictionary<string, List<Dictionary<string, string>>> updatesCache = new Dictionary<string, List<Dictionary<string, string>>>();

    public WinAppReaderImplementation()
    {
        listWithPathsLocalMachine.Add("SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
        listWithPathsLocalMachine.Add("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
        listWithPathsCurrentUser.Add("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall");

        CacheApplicationUpdates();
    }

    private void CacheApplicationUpdates()
    {
        try
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_QuickFixEngineering");
            ManagementObjectCollection collection = searcher.Get();

            foreach (ManagementObject mo in collection)
            {
                string description = (string)mo["Description"];
                string hotFixID = (string)mo["HotFixID"];
                string installedOn = (string)mo["InstalledOn"];

                if (!string.IsNullOrEmpty(description))
                {
                    if (!updatesCache.ContainsKey(description))
                    {
                        updatesCache[description] = new List<Dictionary<string, string>>();
                    }

                    updatesCache[description].Add(new Dictionary<string, string>
                    {
                        { "Description", description },
                        { "HotFixID", hotFixID },
                        { "InstalledOn", installedOn }
                    });
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while caching updates: " + ex.Message);
        }
    }

    List<Dictionary<string, string>> AppReaderImpl.getApplications()
    {
        List<Dictionary<string, string>> apps = new List<Dictionary<string, string>>();

        foreach (string path in listWithPathsLocalMachine)
        {
            ReadRegistryApps(Registry.LocalMachine, path, apps);
        }

        foreach (string path in listWithPathsCurrentUser)
        {
            ReadRegistryApps(Registry.CurrentUser, path, apps);
        }


        return apps;
    }

    private void ReadRegistryApps(RegistryKey rootKey, string path, List<Dictionary<string, string>> apps)
    {
        using (RegistryKey key = rootKey.OpenSubKey(path, false))
        {
            if (key == null) return;

            foreach (string subkeyName in key.GetSubKeyNames())
            {
                using (RegistryKey subkey = key.OpenSubKey(subkeyName))
                {
                    if (subkey?.GetValue("DisplayName") == null) continue;

                    Dictionary<string, string> app = new Dictionary<string, string>
                    {
                        { "DisplayName", (string)subkey.GetValue("DisplayName") },
                        { "InstallDate", (string)subkey.GetValue("InstallDate") },
                        { "DisplayVersion", (string)subkey.GetValue("DisplayVersion") },
                        { "UpdateID", "" },
                        { "UpdateDescription", "" },
                        { "UpdateInstallDate", "" }
                    };

                    // Fetch updates for this application from cache
                    string displayName = (string)subkey.GetValue("DisplayName");
                    var matchingKey = updatesCache.Keys.FirstOrDefault(k => k.Contains(displayName));
                    if (matchingKey != null)
                    {
                        foreach (var update in updatesCache[matchingKey])
                        {
                            app["UpdateID"] = update["HotFixID"];
                            app["UpdateDescription"] = update["Description"];
                            app["UpdateInstallDate"] = update["InstalledOn"];
                        }
                    }

                    /*app["UpdateID"] = "AA";
                    app["UpdateDescription"] = "AA";
                    app["UpdateInstallDate"] = "AA";*/


                    apps.Add(app);
                }
            }
        }
    }
}
