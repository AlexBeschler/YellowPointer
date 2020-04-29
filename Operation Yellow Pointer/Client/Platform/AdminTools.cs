using System;
using LiteDB;

namespace Operation_Yellow_Pointer.Client.Platform
{
    class AdminTools
    {
        /* 
         * Planned Tools:
         * -> Dump data stored for an entry
         * -> Dump data stored for database
         * -> Manual add of entry
         */
        internal static string Database { private get; set; }

        public static void Menu()
        {
            do
            {
                Console.Clear();
                Console.Title = "Admin Tools";

                Console.WriteLine("1. Delete ticker and entries");
                Console.WriteLine("2. Back");

                var input = Console.ReadLine();
                if (input == null)
                    continue;
                var lower = input.ToLower();
                if (lower == "back")
                    return;
                int.TryParse(input, out var num);
                switch (num)
                {
                    case 1:
                        DeleteTicker();
                        break;
                    case 2:
                        return;
                    default:
                        Console.WriteLine("Invalid input");
                        break;
                }
            } while (true);
        }

        private static void DeleteTicker()
        {
            using (var db = new LiteDatabase(Database))
            {
                var success = false;
                var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                string ticker;
                do
                {
                    Console.WriteLine("Ticker of company to delete:");
                    ticker = Console.ReadLine();
                    switch (ticker)
                    {
                        case null:
                            Console.WriteLine("Invalid ticker");
                            continue;
                        case "back":
                            return;
                    }
                    ticker = ticker.ToUpper();

                    //Find ticker, throw error if not found
                    var result = dbKeyRatio.FindOne(Query.EQ("Ticker", ticker));
                    if (result == null)
                        Console.WriteLine("Invalid ticker");
                    else
                        success = true;
                } while (!success);

                try
                {
                    dbKeyRatio.Delete(Query.EQ("Ticker", ticker));
                    Console.WriteLine(ticker + " successfully deleted!");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                    Console.WriteLine("Ticker could not be deleted.");
                    Console.WriteLine("Please see program's error log for more details");
                }
            }
        }
    }
}
