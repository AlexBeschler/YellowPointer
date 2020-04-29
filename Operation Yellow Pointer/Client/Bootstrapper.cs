using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Operation_Yellow_Pointer.Client.Platform;

namespace Operation_Yellow_Pointer.Client
{
    internal static class Bootstrapper
    {
        private const string Version = "2.0";
        private const string DbVersion = "2.0";
        private static char _seperator;
        private static string _dataPath;

        public static void Main(string[] args)
        {
            _seperator = Utility.DetermineOs() == "windows" ? '\\' : '/';
            var currentDirectory = Assembly.GetExecutingAssembly().Location;
            var rootPath = Path.GetFullPath(Utility.DetermineOs() == "windows"
                ? Path.Combine(currentDirectory, @"..\..\..\") //Windows
                : Path.Combine(currentDirectory, @"../../../"));
            _dataPath = rootPath + "Data" + _seperator;
            //Console.WriteLine(Utilities.Seperator);
            Start(args);
        }

        private static void Start(string[] args)
        {
            Initialize(args);
            CheckForUpdates.Update(Version);
            MainMenu.Menu(Version);
        }

        private static void Initialize(string[] args)
        {
            var websitePath = _dataPath + "Website_Data" + _seperator;
            var helpPath = _dataPath + "Help" + _seperator;
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + _seperator;
            var database = _dataPath + "Database.db";

            //Bootstrap values to platform
            MainMenu.WebPath = websitePath;
            Settings.Database = database;
            Settings.DataPath = _dataPath;
            Settings.ErrorLog = _dataPath + "ErrorLog" + _seperator + "Error.txt";
            Settings.Initialize(Version, DbVersion);
            Help.HelpPath = helpPath;
            Utility.DesktopPath = desktop;
            Utility.DataPath = _dataPath;
            Utility.ErrorLog = _dataPath + "ErrorLog" + _seperator + "Error.txt";
            Utility.Temp = _dataPath + _seperator + "temp" + _seperator;
            UpdateDatabase.Temp = _dataPath + _seperator + "temp" + _seperator;
            UpdateDatabase.Database = database;
            AdminTools.Database = database;
            try
            {
                if (args.Contains("-admin"))
                    Utility.AdminMode = true;
            }
            catch (Exception)
            {
                // ignored
            }

            //Bootstrap values to modules
            Modules.Search.Desktop = desktop;
            Modules.Search.Database = database;
            Modules.Search.Data = _dataPath;
            Modules.AddEntry.Desktop = desktop;
            Modules.AddEntry.Database = database;
            Modules.AddEntry.Data = _dataPath;
            Modules.AddEntry.Temp = _dataPath + _seperator + "temp" + _seperator;
            Modules.AddEntry.Init();
            Modules.EditEntry.Desktop = desktop;
            Modules.EditEntry.Database = database;

            //Check if database exists
            if (!File.Exists(database))
            {
                Console.WriteLine("Database file is missing.");
                Console.WriteLine("Press any key to continue...");
                Utility.WriteToErrorLog("Database file missing");
                Console.ReadKey();
                Environment.Exit(1);
            }

            //Read settings file
            var disableAutoWebPage = false;
            var displayChangeLog = true;

            if (File.Exists(_dataPath + "settings.json"))
            {
                using (var sr = new StreamReader(_dataPath + "settings.json"))
                {
                    var reader = new JsonTextReader(sr);
                    var json = JObject.Load(reader);
                    //Check for Disabling Auto Web Page
                    if (json.GetValue("DisableAutoWebPage").Value<bool>())
                        disableAutoWebPage = true;
                    if (!json.GetValue("DisplayChangeLog").Value<bool>())
                        displayChangeLog = false;
                }
            }
            else
                Console.WriteLine("Settings file not found, running on default settings...");
            //Display changelog
            if (displayChangeLog)
            {
                Console.Clear();
                var log = File.ReadAllLines(_dataPath + "ChangeLogs" + _seperator + "ChangeLog_" + Version + ".txt");
                foreach (var t in log)
                    Console.WriteLine(t);
                Console.WriteLine("Press any key to continue...");
                var jsonFile = File.ReadAllText(_dataPath + "settings.json");
                dynamic json = JsonConvert.DeserializeObject(jsonFile);
                json["DisplayChangeLog"] = false;
                string output = JsonConvert.SerializeObject(json, Formatting.Indented);
                File.WriteAllText(_dataPath + "settings.json", output);
                Console.ReadKey();
            }

            //Launch command webpage if args doesn't disable it
            if (disableAutoWebPage) return;
            Console.WriteLine("Launching web page...");
            if (Utility.DetermineOs() == "macos")
                Process.Start("open", websitePath + "index.html");
            else
                Process.Start(websitePath + "index.html");
            Thread.Sleep(1000);
            Console.Clear();
        }
    }
}
