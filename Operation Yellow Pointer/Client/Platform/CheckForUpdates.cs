using System;
using System.Diagnostics;
using System.Net;
using Newtonsoft.Json;
using System.IO.Compression;
using System.IO;

namespace Operation_Yellow_Pointer.Client.Platform
{
    internal class CheckForUpdates
    {
        private const string UpdateLink = @"https://s3.amazonaws.com/yellowpointer/Yellow_Pointer.zip";
        private const string UpdateIndex = @"https://s3.amazonaws.com/yellowpointer/version.json";

        public static void Update(string installedVersion)
        {
            var currentVersion = GetCurrentVersion();
            if (!NewerVersion(installedVersion, currentVersion)) return;
            var success = false;
            do
            {
                Console.WriteLine("There is a newer version of Yellow Pointer available\nWould you like to download it? (Y/N)");
                var result = Console.ReadLine();
                if (result == null)
                {
                    Console.WriteLine("Invalid response");
                    continue;
                }
                result = result.ToLower();
                switch (result)
                {
                    case "n":
                        success = true;
                        break;
                    case "y":
                        Console.WriteLine("Downloading file...");
                        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        desktop += Utility.DetermineOs() == "windows" ? '\\' : '/';
                        using (var wd = new Utility.WebDownload())
                        {
                            wd.DownloadProgressChanged += OnDownloadProgressChanged;
                            
                            wd.DownloadFile(UpdateLink, desktop + "Yellow_Pointer.zip");
                        }
                        //Unzip file
                        Directory.CreateDirectory(desktop + "Yellow_Pointer");
                        ZipFile.ExtractToDirectory(desktop + "Yellow_Pointer.zip", desktop + "Yellow_Pointer");
                        //Delete zip
                        File.Delete(desktop + "Yellow_Pointer.zip");
                        //Run installer
                        if (Utility.DetermineOs() == "macos")
                        {
                            var command = @"mono " + desktop + @"Yellow_Pointer/Yellow\ Pointer\ Installer.exe";
                            Utility.ExecuteCommand(command);
                        }
                        else
                        {
                            Process.Start(desktop + "Yellow_Pointer/Yellow Pointer Installer.exe");
                        }
                        Environment.Exit(0);
                        break;
                    default:
                        Console.WriteLine("Invalid response");
                        break;
                } 
            } while (!success);
            
        }

        private static void OnDownloadProgressChanged(object sender, DownloadProgressChangedEventArgs args)
        {
            var downloaded = args.BytesReceived / Math.Pow(1024, 2);
            var total = args.TotalBytesToReceive / Math.Pow(1024, 2);
            Console.Write("\r" + Math.Round(downloaded, 2) + "/" + Math.Round(total, 2) + " MB");
        }

        private static string GetCurrentVersion()
        {
            try
            {
                string currentVersion;
                using (var wc = new Utility.WebDownload())
                {
                    var tmp = wc.DownloadString(UpdateIndex);
                    dynamic json = JsonConvert.DeserializeObject(tmp);
                    currentVersion = json.version;
                }
                return currentVersion;
            }
            catch (Exception)
            {
                // ignored
            }
            return null;
        }

        private static bool NewerVersion(string installedVersion, string currentVersion)
        {
            var installed = installedVersion.Split('.');
            if (currentVersion == null) return false;
            var current = currentVersion.Split('.');
            for (var i = 0; i < current.Length; i++)
            {
                try
                {
                    if (int.Parse(current[i]) > int.Parse(installed[i]))
                    {
                        return true;
                    }
                    if (int.Parse(current[i]) < int.Parse(installed[i]))
                    {
                        return false;
                    }
                }
                catch (Exception)
                {
                    // ignored
                }
            }
            return false;
        }

        private static void Update()
        {
            //
        }
    }
}
