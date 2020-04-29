using System;
using System.Collections.Generic;
using System.IO;
using LiteDB;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Operation_Yellow_Pointer.Client.Platform
{
    internal class Settings
    {
        internal static string DataPath { private get; set; }
        internal static string ErrorLog { private get; set; }
        internal static string Database { private get; set; }

        public static bool IncludeCompanyInfo;
        public static int MultipleSearchSpacing;
        public static bool FreezeHeaderPane;
        public static bool FreezeTickerColumn;
        public static bool AutoLaunchExcel;

        private static string _version;
        private static string _dbVersion;

        public static void Initialize(string version, string dbVersion)
        {
            _version = version;
            _dbVersion = dbVersion;
            using (var sr = new StreamReader(DataPath + "settings.json"))
            {
                var reader = new JsonTextReader(sr);
                var json = JObject.Load(reader);
                IncludeCompanyInfo = json.GetValue("IncludeCompanyInfo").Value<bool>();
                MultipleSearchSpacing = json.GetValue("BatchSearchSpacing").Value<int>();
                FreezeHeaderPane = json.GetValue("FreezeHeaderRow").Value<bool>();
                FreezeTickerColumn = json.GetValue("FreezeTickerColumn").Value<bool>();
                AutoLaunchExcel = json.GetValue("AutoLaunchExcel").Value<bool>();
            }
        }

        public static void Menu()
        {
            do
            {
                Console.Title = "Settings";
                Console.Clear();
                Console.WriteLine("==Settings==");
                Console.WriteLine("1. Change Settings");
                Console.WriteLine("2. Upload Error Logs");
                Console.WriteLine("3. Restore Backups");
                Console.WriteLine("4. Update Database");
                Console.WriteLine("5. Update Historical Prices");
                Console.WriteLine("6. About Program");
                Console.WriteLine("7. Back");
                var input = Console.ReadLine();
                switch (input)
                {
                    case null:
                        Console.WriteLine("Invalid settings number");
                        continue;
                    case "back":
                        return;
                }
                if (!int.TryParse(input, out var inputNumber))
                {
                    Console.WriteLine("Invalid settings number");
                    continue;
                }

                switch (inputNumber)
                {
                    case 7:
                        return;
                    case 6:
                        int count;
                        using (var db = new LiteDatabase(Database))
                        {
                            var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                            count = dbKeyRatio.Count();
                        }
                        Console.Clear();
                        Console.WriteLine("Yellow Pointer, version " + _version);
                        Console.WriteLine("  Database version: " + _dbVersion);
                        Console.WriteLine("  Number of entries: " + count);
                        Console.WriteLine();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case 5:
                        //
                        break;
                    case 4:
                        Utility.BackUpDatabase();
                        UpdateDatabase.Update();
                        break;
                    case 3:
                        RestoreBackups();
                        break;
                    case 2:
                        Console.WriteLine("Uploading...");
                        var backup = new BackupUtil(DataPath);
                        if (!backup.UploadErrorLogs())
                            Console.ReadKey();
                        break;
                    case 1:
                        ChangeSettings();
                        break;
                    default:
                        Console.WriteLine("Invalid settings number");
                        break;
                }
            } while (true);
        }

        private static void ChangeSettings()
        {
            Console.Title = "Settings > Change Settings";
            do
            {
                Console.Clear();

                Console.WriteLine("==Settings==");
                Console.WriteLine("To change a setting, type the number of the setting. Then " +
                                  "type the value you would like to change it to.");

                Console.Write("1. Prevent Automatic Launch Commands Page: ");
                Console.WriteLine(GetValue("DisableAutoWebPage"));

                Console.Write("2. Include Full Company Info: ");
                Console.WriteLine(GetValue("IncludeCompanyInfo"));

                Console.Write("3. Batch Search Line Spacing: ");
                Console.WriteLine(GetValue("BatchSearchSpacing"));

                Console.Write("4. Freeze Header Row: ");
                Console.WriteLine(GetValue("FreezeHeaderRow"));

                Console.Write("5. Freeze Ticker Column: ");
                Console.WriteLine(GetValue("FreezeTickerColumn"));

                Console.Write("6. Automatically Launch Generated Workbook: ");
                Console.WriteLine(GetValue("AutoLaunchExcel"));

                Console.WriteLine("7. Back");
                var input = Console.ReadLine();
                switch (input)
                {
                    case null:
                        Console.WriteLine("Invalid settings number");
                        continue;
                    case "back":
                        return;
                }
                if (!int.TryParse(input, out var inputNumber))
                {
                    Console.WriteLine("Invalid settings number");
                    continue;
                }

                switch (inputNumber)
                {
                    case 7:
                        return;
                    case 3:
                        var successNum = false;
                        do
                        {
                            Console.WriteLine("Change to: ");
                            var inputNum = Console.ReadLine();
                            if (inputNum == null)
                            {
                                Console.WriteLine("Invalid settings option");
                                continue;
                            }
                            int.TryParse(inputNum, out var result);
                            if (result < 0 || result > 3)
                            {
                                Console.WriteLine("Too much/little line spacing");
                            }
                            else
                            {
                                SetValue("BatchSearchSpacing", result);
                                MultipleSearchSpacing = result;
                                successNum = true;
                            }
                        } while (!successNum);
                        break;
                    case 6:
                    case 5:
                    case 4:
                    case 2:
                    case 1:
                        var successOption = false;
                        string setting;
                        do
                        {
                            Console.WriteLine("Change to:");
                            setting = Console.ReadLine();
                            if (setting == null)
                            {
                                Console.WriteLine("Invalid settings option");
                                continue;
                            }
                            setting = setting.ToLower();
                            if (setting != "enable" && setting != "disable")
                            {
                                Console.WriteLine("Invalid settings option");
                                continue;
                            }
                            successOption = true;
                        } while (!successOption);

                        var value = setting == "enable";
                        switch (inputNumber)
                        {
                            case 1:
                                SetValue("DisableAutoWebPage", value);
                                break;
                            case 2:
                                SetValue("IncludeCompanyInfo", value);
                                IncludeCompanyInfo = value;
                                break;
                            case 4:
                                SetValue("FreezeHeaderRow", value);
                                FreezeHeaderPane = value;
                                break;
                            case 5:
                                SetValue("FreezeTickerColumn", value);
                                FreezeTickerColumn = value;
                                break;
                            case 6:
                                SetValue("AutoLaunchExcel", value);
                                AutoLaunchExcel = value;
                                break;
                            default:
                                Console.WriteLine("Invalid settings number");
                                continue;
                        }
                        break;
                    default:
                        Console.WriteLine("Invalid settings number");
                        break;
                }
            } while (true);
        }

        private static string GetValue(string key)
        {
            using (var sr = new StreamReader(DataPath + "settings.json"))
            {
                var reader = new JsonTextReader(sr);
                var json = JObject.Load(reader);
                var result = json.GetValue(key).ToString();
                //Sort if value is bool or int
                if (result == "True" || result == "False")
                {
                    return result == "True" ? "Enabled" : "Disabled";
                }
                return int.TryParse(result, out var num) ? num.ToString() : null;
            }
        }

        private static void SetValue(string key, bool value)
        {
            var jsonFile = File.ReadAllText(DataPath + "settings.json");
            dynamic json = JsonConvert.DeserializeObject(jsonFile);
            json[key] = value;
            string output = JsonConvert.SerializeObject(json, Formatting.Indented);
            File.WriteAllText(DataPath + "settings.json", output);
        }

        private static void SetValue(string key, int value)
        {
            var jsonFile = File.ReadAllText(DataPath + "settings.json");
            dynamic json = JsonConvert.DeserializeObject(jsonFile);
            json[key] = value;
            string output = JsonConvert.SerializeObject(json, Formatting.Indented);
            File.WriteAllText(DataPath + "settings.json", output);
        }

        private static void RestoreBackups()
        {
            var backup = new BackupUtil(DataPath);
            var list = backup.GetIndexListing() ?? new List<BackupUtil.Index>();

            var listing = 1;
            Console.WriteLine("Listing |        Date         | No. of Entries");
            Console.WriteLine("--------|---------------------|---------------");
            for (var i = list.Count - 1; i >= 0; i--)
            {
                Console.WriteLine(listing + ".      | " + list[i].Date + " | " + list[i].DatabaseEntries);
                listing++;
            }
            Console.WriteLine();
            Console.WriteLine("Which backup would you like to restore?");

            var success = false;
            do
            {
                var input = Console.ReadLine();
                switch (input)
                {
                    case null:
                        Console.WriteLine("Invalid backup");
                        continue;
                    case "back":
                        return;
                }
                if (!int.TryParse(input, out var inputNumber))
                {
                    Console.WriteLine("Invalid backup");
                    continue;
                }
                inputNumber = list.Count - inputNumber;
                if (inputNumber < 0 || inputNumber >= list.Count)
                {
                    Console.WriteLine("Invalid backup");
                    continue;
                }
                backup.RestoreBackup(list[inputNumber].Guid);
                Console.WriteLine();
                success = true;
            } while (!success);
            Console.WriteLine("Database restoration successful!");
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }
    }
}
