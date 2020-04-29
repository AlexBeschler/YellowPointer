using System;
using System.Diagnostics;
using Operation_Yellow_Pointer.Client.Modules;

namespace Operation_Yellow_Pointer.Client.Platform
{
    class MainMenu
    {
        internal static string WebPath { private get; set; }

        public static void Menu(string version)
        {
            do
            {
                Console.Clear();
                Console.Title = "Yellow Pointer v" + version;
                if (Utility.AdminMode)
                    Console.Title += " (Admin Mode)";
                var title = "===Yellow Pointer Main Menu, v" + version;
                if (Utility.AdminMode)
                    title += " (Admin Mode)";
                Console.WriteLine(title + "===");
                Console.WriteLine("1. Search Database");
                Console.WriteLine("2. Add Database Entries");
                Console.WriteLine("3. Edit Database Entries");
                Console.WriteLine("4. Launch Commands Web Page");
                Console.WriteLine("5. Settings");
                Console.WriteLine("6. Help");
                Console.WriteLine("7. Exit");
                if (Utility.AdminMode)
                    Console.WriteLine("8. Admin Tools");

                var input = Console.ReadLine();
                if (input == null)
                    continue;
                var lower = input.ToLower();
                if (lower == "exit")
                    Environment.Exit(0);
                int.TryParse(input, out var num);
                switch (num)
                {
                    case 1:
                        Search.Menu();
                        break;
                    case 2:
                        Utility.BackUpDatabase();
                        AddEntry.Menu();
                        break;
                    case 3:
                        Utility.BackUpDatabase();
                        EditEntry.Menu();
                        break;
                    case 4:
                        if (Utility.DetermineOs() == "macos")
                            Process.Start("open", WebPath + "index.html");
                        else
                            Process.Start(WebPath + "index.html");
                        break;
                    case 5:
                        Settings.Menu();
                        break;
                    case 6:
                        Help.DisplayHelp();
                        break;
                    case 7:
                        Environment.Exit(0);
                        break;
                    case 8:
                        if (!Utility.AdminMode)
                        {
                            Console.WriteLine("Invalid input");
                            break;
                        }
                        AdminTools.Menu();
                        break;
                    default:
                        Console.WriteLine("Invalid input");
                        break;
                }
            } while (true);
        }
    }
}
