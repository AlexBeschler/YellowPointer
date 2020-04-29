using System;
using System.IO;
using System.Timers;
using LiteDB;
using OfficeOpenXml;
using Operation_Yellow_Pointer.Client.Platform;

namespace Operation_Yellow_Pointer.Client.Modules
{
    internal static class EditEntry
    {
        internal static string Database;
        internal static string Desktop;

        private static int _line;
        private static int _total;

        private const string ErrorLog = "EditDatabaseErrorList.txt";

        public static void Menu()
        {
            Console.Title = "Edit Database Entries";
            do
            {
                Console.Clear();
                Console.WriteLine("Edit Database Entries:");
                Console.WriteLine("1. Import classification revisions from Excel");
                Console.WriteLine("2. Revise specific classification");
                Console.WriteLine("3. Add CIK to ticker");
                Console.WriteLine("4. Back");
                var input = Console.ReadLine();
                switch (input)
                {
                    case null:
                        Console.WriteLine("Invalid input");
                        continue;
                    case "back":
                        return;
                    case "1":
                        Batch();
                        break;
                    case "2":
                        Manual();
                        break;
                    case "3":
                        AddCik();
                        break;
                    case "4":
                        return;
                }
            } while (true);
        }

        private static void Batch()
        {
            var success = false;
            string input;
            do
            {
                Console.WriteLine("Path to Excel file:");
                input = Console.ReadLine();
                switch (input)
                {
                    case null:
                        Console.WriteLine("Invalid input");
                        continue;
                    case "back":
                        return;
                }
                input = input.Replace("\"", string.Empty);
                input = input.Replace(" ", string.Empty);
                if (!File.Exists(input))
                {
                    Console.WriteLine("Could not find Excel file. Please try again.");
                }
                else
                {
                    success = true;
                }
            } while (!success);
            var file = new FileInfo(input);
            Utility.CheckIfFileIsLocked(file);
            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets[1];
                _total = worksheet.Dimension.End.Row;
                var timer = new Timer(1000);
                timer.Elapsed += UpdateProgress;
                Console.Write("Progress: 0.00%");
                timer.Enabled = true;
                _line = 1;
                var edited = 0;

                using (var db = new LiteDatabase(Database))
                {
                    var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                    for (; _line <= _total; _line++)
                    {
                        var update = false;
                        var ticker = worksheet.Cells[_line, 1].Value.ToString();
                        var result = dbKeyRatio.FindOne(Query.EQ("Ticker", ticker));
                        if (result == null)
                        {
                            Utility.WriteToLog(ticker, "Could not find ticker", ErrorLog);
                            continue;
                        }

                        var sector = worksheet.Cells[_line, 2].Value;
                        if (sector != null)
                        {
                            if (sector.ToString() != "" && sector.ToString() != string.Empty)
                            {
                                result.Sector = sector.ToString();
                                update = true;
                            }
                        }
                        
                        var industry = worksheet.Cells[_line, 3].Value;
                        if (industry != null)
                        {
                            if (industry.ToString() != "" && industry.ToString() != string.Empty)
                            {
                                result.Industry = industry.ToString();
                                update = true;
                            }
                        }
                        
                        var subindustry = worksheet.Cells[_line, 4].Value;
                        if (subindustry != null)
                        {
                            if (subindustry.ToString() != "" && subindustry.ToString() != string.Empty)
                            {
                                result.Subindustry = subindustry.ToString();
                                update = true;
                            }
                        }

                        if (!update) continue;
                        try
                        {
                            dbKeyRatio.Update(result);
                            edited++;
                        }
                        catch (Exception e)
                        {
                            Utility.WriteToErrorLog(e);
                        }
                    }
                }
                timer.Enabled = false;
                Console.WriteLine();
                if (!File.Exists(Desktop + "EditDatabaseErrorList.txt"))
                    Console.WriteLine("Success!");
                if (File.Exists(Desktop + "EditDatabaseErrorList.txt"))
                    Console.WriteLine("Some tickers could not be edited.\nSee error on desktop.");
                Console.WriteLine("Edited " + edited + " tickers.");
                Console.WriteLine("Press any key to continue");
                Console.ReadKey();
            }
        }

        private static void Manual()
        {
            using (var db = new LiteDatabase(Database))
            {
                var success = false;
                string metric;
                var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                KeyRatio result = null;

                do
                {
                    Console.WriteLine("Ticker of company to edit:");
                    var ticker = Console.ReadLine();
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
                    result = dbKeyRatio.FindOne(Query.EQ("Ticker", ticker));
                    if (result == null)
                        Console.WriteLine("Invalid ticker");
                    else
                        success = true;
                } while (!success);

                success = false;
                do
                {
                    Console.WriteLine("Metric to edit:");
                    metric = Console.ReadLine();
                    switch (metric)
                    {
                        case null:
                            Console.WriteLine("Invalid metric");
                            continue;
                        case "back":
                            return;
                    }
                    metric = metric.ToLower();

                    if (metric == "sector" || metric == "industry" || metric == "subindustry")
                    {
                        success = true;
                    }
                    else
                    {
                        Console.WriteLine("Invalid metric");
                    }
                } while (!success);

                Console.WriteLine("New value:");
                var newValue = Console.ReadLine();
                if (newValue == "back")
                {
                    return;
                }
                switch (metric)
                {
                    case "sector":
                        result.Sector = newValue;
                        break;
                    case "industry":
                        result.Industry = newValue;
                        break;
                    case "subindustry":
                        result.Subindustry = newValue;
                        break;
                }
                try
                {
                    dbKeyRatio.Update(result);
                    Console.WriteLine("Database entry written successfully!");
                    Console.WriteLine(
                        "Company " +
                        result.Name +
                        " (" +
                        result.Ticker +
                        ")'s " +
                        metric +
                        " is now '" +
                        newValue +
                        "'");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                    Console.WriteLine("Ticker could not be edited.");
                    Console.WriteLine("Please see program's error log for more details");
                }
            }
        }

        private static void UpdateProgress(object source, ElapsedEventArgs e)
        {
            //(Actual Work / Work) * 100
            var progress = _line / (double)_total;
            progress *= 100.00;
            Console.Write("\r                ");
            Console.Write("\rProgress: " + Math.Round(progress, 2) + "%");
        }

        private static void AddCik()
        {
            using (var db = new LiteDatabase(Database))
            {
                var success = false;
                string cik;
                var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                KeyRatio result = null;

                do
                {
                    Console.WriteLine("Ticker of company to add CIK:");
                    var ticker = Console.ReadLine();
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
                    result = dbKeyRatio.FindOne(Query.EQ("Ticker", ticker));
                    if (result == null)
                        Console.WriteLine("Invalid ticker");
                    else
                        success = true;
                } while (!success);

                success = false;
                do
                {
                    Console.WriteLine("CIK (please include beginning zeros!): ");
                    cik = Console.ReadLine();
                    switch (cik)
                    {
                        case null:
                            Console.WriteLine("Invalid metric");
                            continue;
                        case "back":
                            return;
                        default:
                            success = true;
                            break;
                    }
                } while (!success);

                result.Cik = cik;

                try
                {
                    dbKeyRatio.Update(result);
                    Console.WriteLine("CIK updated successfully!");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                    Console.WriteLine("Ticker's CIK could not be edited.");
                    Console.WriteLine("Please see program's error log for more details");
                }
            }
        }
    }
}
