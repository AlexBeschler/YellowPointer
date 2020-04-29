using System;
using System.IO;

namespace Operation_Yellow_Pointer.Client.Platform
{
    internal class Help
    {
        internal static string HelpPath { private get; set; }

        internal static void DisplayHelp()
        {
            var main = HelpPath + "Main.txt";
            var batchOperation = HelpPath + "BatchOperations.txt";
            var analytics = HelpPath + "Analytics.txt";

            Console.Title = "Help";
            do
            {
                Console.Clear();
                var f = File.ReadAllLines(main);
                foreach (var s in f)
                {
                    Console.WriteLine(s);
                }
                var input = Console.ReadLine();
                if (input == null)
                {
                    Console.WriteLine("Invalid input");
                    continue;
                }
                input = input.ToLower();
                switch (input)
                {
                    case "batch":
                        Display(batchOperation);
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "back":
                        return;
                    default:
                        Console.WriteLine("Invalid input");
                        break;
                }
            } while (true);
        }

        private static void Display(string file)
        {
            var f = File.ReadAllLines(file);
            foreach (var s in f)
            {
                Console.WriteLine(s);
            }
        }
    }
}
