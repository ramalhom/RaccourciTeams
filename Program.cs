using System;
using System.Text;
using Microsoft.SharePoint.Client;
using IWshRuntimeLibrary;
using System.IO;
using System.Diagnostics;
using System.Configuration;
using System.Collections.Specialized;

namespace RaccourciTeams
{
    class Program
    {
        // URL of the Teams SharePoint Base
        static string URLTeamSitesBase;
        // Domains 
        static string[] Domains;
        // Network URL of the Teams SharePoint Base
        static string NetworkTeamSitesBase;
        // Verbose information
        static bool Verbose = false;
        // Name of Base Folder
        static string NameBaseFolder;
        // Document Library in SharePoint 
        static string DocsLibrary;
        // Counter for displaying the number of shortcut connected
        static int Count = 0;

        static void Main(string[] args)
        {
            // Get Variables for App.config
            URLTeamSitesBase = ConfigurationManager.AppSettings.Get("URLTeamSitesBase");
            Domains = ConfigurationManager.AppSettings.Get("DomainsTeamSites").Split(',');
            Verbose = ConfigurationManager.AppSettings.Get("Verbose").ToLower()=="true"?true:false;
            NameBaseFolder = ConfigurationManager.AppSettings.Get("NameBaseFolder");
            DocsLibrary = ConfigurationManager.AppSettings.Get("DocsLibrary");
            NetworkTeamSitesBase = URLTeamSitesBase.Replace("https://", "\\\\").Replace("/", "\\");

            if (Verbose)
            {
                Console.WriteLine("************************************************");
                Console.WriteLine("Connexion des raccourcis vers les sites d'équipe");
                Console.WriteLine("************************************************");
            }

            // Delete TEAMS Folder and all shortcuts
            RemoveAllShortcut();
            // Create base directory
            CreateDirectoryBase();
            // Connect to SharePoint Teams (NET USE)
            ConnectWebs(NetworkTeamSitesBase, Domains);
            // Get SubWeb that user have access and create Shortcut
            GetSubWebsAndCreateShortcut(URLTeamSitesBase, NetworkTeamSitesBase, Domains);
            if (Verbose)
            {
                Console.WriteLine("Nombre de raccourci créé : " + Count);
                System.Threading.Thread.Sleep(2000);
            }
        }

        public static void ConnectWebs(string path, string[] domains)
        {
            try
            {
                // For each Domain connect to the share with net use
                foreach (string domain in domains)
                {
                    string newpath = path + domain;
                    Process process = new Process();
                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    startInfo.FileName = "cmd.exe";
                    startInfo.Arguments = "/C net use "+ newpath;
                    process.StartInfo = startInfo;
                    Process.Start(startInfo);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        public static void GetSubWebsAndCreateShortcut(string path,string shortcutpath, string[] domains)
        {
            try
            {
                // Foreach Domain Get SubWeb For Current User
                foreach (string domain in domains)
                {
                    ClientContext clientContext = new ClientContext(path + domain);
                    WebCollection oWebsite = clientContext.Web.GetSubwebsForCurrentUser(null);
                    clientContext.Load(oWebsite);
                    clientContext.ExecuteQuery();
                    foreach (Web orWebsite in oWebsite)
                    {
                        string newpath = shortcutpath + orWebsite.ServerRelativeUrl.Substring(1).Replace("/","\\");
                        CreateShortcut(orWebsite.Title, newpath, path + orWebsite.ServerRelativeUrl.Substring(1));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public static void CreateShortcut(string name, string path, string url)
        {

            string BaseDirectory = Environment.GetFolderPath(Environment.SpecialFolder.NetworkShortcuts) + "\\" + NameBaseFolder;
            
            // Create the shortcut directory
            Directory.CreateDirectory(BaseDirectory + "\\" + name);
            
            // Set the folder as read only
            (new DirectoryInfo(BaseDirectory + "\\" + name)).Attributes = FileAttributes.ReadOnly;

            // Create the ini file
            if (!System.IO.File.Exists(BaseDirectory + "\\" + name + "\\desktop.ini"))
            {
                using (StreamWriter sw = System.IO.File.CreateText(BaseDirectory + "\\" + name + "\\desktop.ini"))
                {
                    sw.WriteLine("[.ShellClassInfo]");
                    sw.WriteLine("CLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}");
                    sw.WriteLine("Flags=2");
                }
            }
            System.IO.File.SetAttributes(BaseDirectory + "\\" + name + "\\desktop.ini", FileAttributes.Hidden| FileAttributes.System);

            // Create the shortcut
            IWshShortcut shortcut;
            
            var wshShell = new WshShell();
            shortcut = (IWshShortcut)wshShell.CreateShortcut(Path.Combine(BaseDirectory + "\\" + name, "target.lnk"));

            string targetpath = path.Replace(".ch\\",".ch@SSL\\") + "\\" + DocsLibrary;
            shortcut.TargetPath = targetpath;
            shortcut.Description = targetpath;
            shortcut.RelativePath = targetpath;
            shortcut.WorkingDirectory = targetpath;
            Count++;
            if (Verbose) Console.WriteLine(name);
            shortcut.Save();
        }

        public static void CreateDirectoryBase()
        {
            string BaseDirectory = Environment.GetFolderPath(Environment.SpecialFolder.NetworkShortcuts) + "\\" + NameBaseFolder;
            
            // Create the base directory (TEAMS)
            Directory.CreateDirectory(BaseDirectory);

            var iniPath = Path.Combine(BaseDirectory, "desktop.ini");
            // Create new ini file with the required contents
            var iniContents = new StringBuilder()
                .AppendLine("[.ShellClassInfo]")
                .AppendLine($"IconResource=C:\\WINDOWS\\System32\\SHELL32.dll,13")
                .ToString();
            System.IO.File.WriteAllText(iniPath, iniContents);
            // Hide the ini file and set it as system
            System.IO.File.SetAttributes(
               iniPath,
               System.IO.File.GetAttributes(iniPath) | FileAttributes.Hidden | FileAttributes.System);
            // Set the folder as system
            System.IO.File.SetAttributes(
                BaseDirectory,
                System.IO.File.GetAttributes(BaseDirectory) | FileAttributes.System);
        }

        public static void RemoveAllShortcut()
        {
            string BaseDirectory = Environment.GetFolderPath(Environment.SpecialFolder.NetworkShortcuts) + "\\" + NameBaseFolder;
            if (Directory.Exists(BaseDirectory))
            {
                // Remove all shortcuts on base directory
                foreach (string teamdir in Directory.GetDirectories(BaseDirectory))
                {
                    string[] filesteamdir = Directory.GetFiles(teamdir, "*");
                    foreach (string file in filesteamdir)
                    {
                        System.IO.File.Delete(file);
                    }
                    (new DirectoryInfo(teamdir)).Attributes = FileAttributes.Normal;
                    Directory.Delete(teamdir);
                }
                // Remove the base directory with the files
                string[] filesteamroot = Directory.GetFiles(BaseDirectory, "*");
                foreach (string file in filesteamroot)
                {
                    System.IO.File.Delete(file);
                }
                (new DirectoryInfo(BaseDirectory)).Attributes = FileAttributes.Normal;
                Directory.Delete(BaseDirectory);
            }
        }
    }
}
