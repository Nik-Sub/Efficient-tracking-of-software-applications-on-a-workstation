using System.Diagnostics;
using Microsoft.Win32;


class WinAppReaderImplementation : AppReaderImpl
{

    List<string> listWithPathsCurrentUser = new List<string>();
    List<string> listWithPathsLocalMachine = new List<string>();

    
    public WinAppReaderImplementation(){
        //listWithPaths.Add("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths");
        //listWithPaths.Add("SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths");
        listWithPathsLocalMachine.Add("SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
        listWithPathsLocalMachine.Add("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
        listWithPathsCurrentUser.Add("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
    }

    List<Dictionary<string, string>> AppReaderImpl.getApplications()
    {
        List<Dictionary<string, string>> apps = new List<Dictionary<string, string>>();
        foreach(string path in listWithPathsLocalMachine){
            // Get the 64-bit installed applications
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(path, true))
            {
                foreach (string subkeyName in key.GetSubKeyNames())
                {
                    Dictionary<string, string> app = new Dictionary<string, string>();

                    using (RegistryKey subkey = key.OpenSubKey(subkeyName))
                    {
                        // Console.WriteLine(subkey.ToString());
                        // Console.WriteLine(subkey.GetValue("DisplayName") + " " + subkey.GetValue("DisplayVersion")
                        //                + " " + subkey.GetValue("InstallDate"));
                        if (subkey.GetValue("DisplayName") == null){
                            //Console.WriteLine(subkeyName);
                            continue;
                        }
                        //Console.WriteLine(path);
                        app.Add("DisplayName", (string)subkey.GetValue("DisplayName"));
                        app.Add("InstallDate", (string)subkey.GetValue("InstallDate"));
                        app.Add("DisplayVersion", (string)subkey.GetValue("DisplayVersion"));
                        app.Add("InstallLocation", (string)subkey.GetValue("InstallLocation"));
                        //Console.WriteLine(app["DisplayName"]);
                    }

                    apps.Add(app);
                }
            }
        }


        foreach(string path in listWithPathsCurrentUser){
            // Get the 64-bit installed applications
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path, true))
            {
                foreach (string subkeyName in key.GetSubKeyNames())
                {
                    Dictionary<string, string> app = new Dictionary<string, string>();

                    using (RegistryKey subkey = key.OpenSubKey(subkeyName))
                    {
                        // Console.WriteLine(subkey.ToString());
                        // Console.WriteLine(subkey.GetValue("DisplayName") + " " + subkey.GetValue("DisplayVersion")
                        //                + " " + subkey.GetValue("InstallDate"));
                        if (subkey.GetValue("DisplayName") == null){
                            //Console.WriteLine(subkeyName);
                            continue;
                        }
                        //Console.WriteLine(path);
                        app.Add("DisplayName", (string)subkey.GetValue("DisplayName"));
                        app.Add("InstallDate", (string)subkey.GetValue("InstallDate"));
                        app.Add("DisplayVersion", (string)subkey.GetValue("DisplayVersion"));
                        app.Add("InstallLocation", (string)subkey.GetValue("InstallLocation"));
                        //Console.WriteLine(app["DisplayName"]);
                    }

                    apps.Add(app);
                }
            }
        }

        

        return apps;
    }
}