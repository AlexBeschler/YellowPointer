using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using LiteDB;
using OfficeOpenXml;
using Operation_Yellow_Pointer.Client.Platform;

namespace Operation_Yellow_Pointer.Client.Modules
{
    internal static class Search
    {
        internal static string Database;
        internal static string Desktop;
        internal static string Data;

        private static string[] _key;
        private static string[] _value;

        private static string ErrorLog = "SearchQueryError.txt";

        public static void Menu()
        {
            _key = File.ReadAllLines(Data + "Key.txt");
            _value = File.ReadAllLines(Data + "Value.txt");

            Console.Title = "Search Database";
            using (var db = new LiteDatabase(Database))
            {
                var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");

                do
                {
                    Console.Clear();
                    Console.WriteLine("Search Database:");
                    Console.WriteLine("1. Get general information about a company");
                    Console.WriteLine("2. Get specific metric info");
                    Console.WriteLine("3. Import query from Excel");
                    Console.WriteLine("4. Export tickers and classifications");
                    var input = Console.ReadLine();
                    switch (input)
                    {
                        case null:
                            Console.WriteLine("Invalid response");
                            break;
                        case "back":
                            return;
                    }
                    if (input != "1" && input != "2" && input != "3" && input != "4")
                    {
                        Console.WriteLine("Invalid response");
                        continue;
                    }

                    switch (input)
                    {
                        case "1":
                            GeneralInfo(dbKeyRatio);
                            break;
                        case "2":
                            SpecificInfo(dbKeyRatio);
                            break;
                        case "3":
                            ImportQuery(dbKeyRatio);
                            break;
                        case "4":
                            ExportClassifications(dbKeyRatio);
                            break;
                    }
                } while (true);
            }
        }

        #region MenuMethods

        private static void GeneralInfo(LiteCollection<KeyRatio> dbKeyRatio)
        {
            const string bullet = "  -> ";
            do
            {
                Console.WriteLine("Search by ticker:");
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
                var result = dbKeyRatio.FindOne(Query.EQ("Ticker", ticker));
                if (result == null)
                {
                    Console.WriteLine("Could not find ticker");
                }
                else
                {
                    Console.Clear();
                    Console.WriteLine(result.Name + ":");
                    Console.WriteLine(bullet + "Ticker: " + result.Ticker + ", CIK: " + result.Cik);
                    Console.WriteLine(bullet + "Currency Type: " + result.CurrencyType);
                    Console.WriteLine(bullet + "Fiscal Year Ends: " + result.FiscalYearEnd);
                    Console.WriteLine();
                    Console.WriteLine(bullet + result.Ticker + " belongs to the following catagories: ");
                    Console.WriteLine("  " + bullet + "Sector: " + result.Sector);
                    Console.WriteLine("  " + bullet + "Industry: " + result.Industry);
                    Console.WriteLine("  " + bullet + "Subindustry: " + result.Subindustry);
                    Console.WriteLine();
                    Console.WriteLine(bullet + 
                        "The database contains " + 
                        result.Ticker + 
                        "'s financial data from " + 
                        GetYear(result.FiscalYears[0]) + 
                        " to " + 
                        GetYear(result.FiscalYears[result.FiscalYears.Length - 1]));
                    Console.WriteLine();
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    return;
                }
            } while (true);
        }

        private static void SpecificInfo(LiteCollection<KeyRatio> dbKeyRatio)
        {
            //TODO: Add ability to convert to USD if foreign currency
            var metrics = new List<string>();
            string[] years = { };
            var results = new List<KeyRatio>();
            Console.WriteLine("Search by ticker, sector, industry, subindustry");
            Console.WriteLine("Start command with parameter to search with:");
            Console.WriteLine("    i.e. 'ticker | a, b, c'; 'sector | media, healthcare'; etc.");

            /* Command Search */
            var valid = false;
            do
            {
                var input = Console.ReadLine();
                switch (input)
                {
                    case null:
                        Console.WriteLine("Invalid search parameters");
                        continue;
                    case "back":
                        return;
                }

                var splice = input.Split('|');
                if (splice.Length != 2)
                {
                    Console.WriteLine("Invalid command");
                    continue;
                }
                var command = splice[0].Trim();
                command = Utility.UppercaseFirst(command);
                var parameters = splice[1].Split(',').ToList();
                if (parameters.Count < 1)
                {
                    Console.WriteLine("Invalid command");
                    continue;
                }
                for (var i = 0; i < parameters.Count; i++)
                {
                    parameters[i] = parameters[i].Trim();
                }

                if (
                    !command.Equals("Ticker") &&
                    !command.Equals("Sector") &&
                    !command.Equals("Industry") &&
                    !command.Equals("Subindustry"))
                {
                    Console.WriteLine("Invalid command\nPossible misspelling?");
                    continue;
                }

                Console.Write("Loading...");
                var unfound = new List<string>();
                foreach (var parameter in parameters)
                {
                    var query = dbKeyRatio.Find(Query.EQ(command, parameter)).ToList();
                    if (query.Count == 0)
                        unfound.Add(parameter);
                    else
                        results.AddRange(query);
                }

                Console.Write("\r");

                if (results.Count < 1)
                {
                    Console.WriteLine("No results found.");
                    continue;
                }

                Console.WriteLine(results.Count == 1 ? "1 result found" : results.Count + " results found");
                if (unfound.Count > 0)
                {
                    Console.Write(unfound.Count == 1 ? "1 unfound search parameter: " : unfound.Count + " unfound search parameters: ");
                    FormatResults(unfound);
                }
                valid = true;
            } while (!valid);

            /* Metric Search */
            var invalidParam = false;
            valid = false;
            do
            {
                Console.WriteLine("Metric:");
                var input = Console.ReadLine();
                switch (input)
                {
                    case null:
                        Console.WriteLine("Invalid metric(s)");
                        continue;
                    case "back":
                        return;
                }
                metrics = input.Split(',').ToList();
                metrics = Utility.RemoveWhiteSpace(metrics);
                //Find if valid metric
                for (var i = 0; i < metrics.Count; i++)
                {
                    metrics[i] = metrics[i].ToUpper();
                    if (_key.Contains(metrics[i])) continue;
                    Console.WriteLine("Invalid metric: " + metrics[i]);
                    invalidParam = true;
                    break;
                }
                if (!invalidParam)
                {
                    valid = true;
                }
            } while (invalidParam || !valid);

            /* Year Search */
            valid = false;
            do
            {
                Console.WriteLine("Year:");
                var readLine = Console.ReadLine();
                switch (readLine)
                {
                    case null:
                        Console.WriteLine("Invalid year(s)");
                        continue;
                    case "back":
                        return;
                }

                //TODO: Deal with years
                years = readLine.Split(',');
                years = Utility.RemoveWhiteSpace(years);
                valid = true;
            } while (!valid);

            var pos = SortYears(results, years);
            Console.WriteLine("Exporting...");
            WriteToExcel(results, metrics, years, pos, true);
        }

        private static void ImportQuery(LiteCollection<KeyRatio> dbKeyRatio)
        {
            //TODO: Allow ability to convert foreign currency
            var success = false;
            string line;
            do
            {
                Console.WriteLine("Path to Excel file:");
                line = Console.ReadLine();
                switch (line)
                {
                    case null:
                        Console.WriteLine("Invalid file");
                        continue;
                    case "back":
                        return;
                }
                line = line.Replace("\"", string.Empty);
                if (!File.Exists(line))
                {
                    Console.WriteLine("Could not find Excel file. Please try again.");
                }
                else
                {
                    success = true;
                }
            } while (!success);
            var file = new FileInfo(line);

            //Used as the package for the generated file
            var newFile = new FileInfo(GetFileName());
            //Create new Excel file
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Data");
                package.SaveAs(newFile);
            }

            Utility.CheckIfFileIsLocked(file);
            
            using (var package = new ExcelPackage(file))
            {
                var header = 1; //This is used to point to which row to write the new information
                //Need to get rows and use as length in the for loop
                var worksheet = package.Workbook.Worksheets[1];
                var height = worksheet.Dimension.End.Row;
                if (height == 1)
                {
                    //Freeze both header pane and ticker column
                    if (Settings.FreezeHeaderPane && Settings.FreezeTickerColumn)
                    {
                        worksheet.View.FreezePanes(2, 2);
                    }
                    //Freeze just the header pane
                    else if (Settings.FreezeHeaderPane && !Settings.FreezeTickerColumn)
                    {
                        worksheet.View.FreezePanes(2, 1);
                    }
                    //Freeze just the ticker column
                    else if (!Settings.FreezeHeaderPane && Settings.FreezeTickerColumn)
                    {
                        worksheet.View.FreezePanes(1, 2);
                    }
                }
                Console.WriteLine("Exporting...");
                for (var i = 1; i <= height; i++) //i indicates how many search queries there were
                {
                    //Repeat until reached end of the worksheet
                    var length = worksheet.Dimension.End.Column;
                    var c = worksheet.Cells[i, 1].Value;
                    if (c == null)
                    {
                        Console.WriteLine("Failed to extract data. See error log for details.");
                        Utility.WriteToLog("Invalid command", ErrorLog);
                        return;
                    }
                    var command = c.ToString();
                    command = command.Trim();
                    command = Utility.UppercaseFirst(command);
                    if (command != "Ticker" &&
                        command != "Sector" &&
                        command != "Industry" &&
                        command != "Subindustry")
                    {
                        Console.WriteLine("Failed to extract data. See error log for details.");
                        Utility.WriteToLog(command + ": Invalid command", ErrorLog);
                        return;
                    }
                    var j = 2;
                    var results = new List<KeyRatio>();
                    while (j < length)
                    {
                        var p = worksheet.Cells[i, j].Value;
                        if (p == null)
                        {
                            j++;
                            continue;
                        }
                        var parameter = p.ToString();
                        if (parameter == "|" && j == 2)
                        {
                            Utility.WriteToLog("Found a | character in column B; ignoring...", ErrorLog);
                            j++;
                        }
                        if (parameter == "|")
                        {
                            j++;
                            break;
                        }
                        var query = dbKeyRatio.Find(Query.EQ(command, parameter)).ToList();
                        if (query.Count == 0)
                        {
                            Console.WriteLine("Unrecognized parameter: " + parameter + "; ignoring...");
                            Utility.WriteToLog(parameter + ": parameter not recognized", ErrorLog);
                        }
                        else
                            results.AddRange(query);
                        j++;
                    }
                    if (results.Count < 1)
                    {
                        Console.WriteLine("Failed to extract data. See error log for details.");
                        Utility.WriteToLog("Not enough parameters", ErrorLog);
                    }
                    //Metrics
                    var m = new List<string>();
                    while (j <= length)
                    {
                        var tmp = worksheet.Cells[i, j].Value;
                        if (tmp == null)
                        {
                            j++;
                            continue;
                        }
                        var cell = tmp.ToString();
                        cell = cell.ToUpper();
                        if (cell == "|")
                        {
                            j++;
                            break;
                        }
                        m.Add(cell);
                        j++;
                    }
                    var metrics = m.ToList();

                    var invalidMetrics = new List<string>();
                    foreach (var t in metrics)
                    {
                        if (_key.Contains(t)) continue;
                        Console.WriteLine("Failed to extract data. See error log for details.");
                        Utility.WriteToLog(t + ": Unrecognized metric", ErrorLog);
                        invalidMetrics.Add(t);
                    }
                    foreach (var invalidMetric in invalidMetrics)
                        metrics.RemoveAt(metrics.IndexOf(invalidMetric));
                    //Years
                    var y = new List<string>();
                    while (j <= length)
                    {
                        var tmp = worksheet.Cells[i, j].Value;
                        if (tmp == null)
                        {
                            j++;
                            continue;
                        }
                        var cell = tmp.ToString();
                        y.Add(cell);
                        j++;
                    }
                    var years = y.ToArray();
                    var sortedYears = SortYears(results, years);
                    
                    header = AppendToExcel(newFile, false, results, metrics, years, sortedYears, header, false);
                }
            }
            Console.WriteLine("Excel file generated successfully!");
            if (Settings.AutoLaunchExcel)
            {
                Console.WriteLine("Launching Excel...");
                if (Utility.DetermineOs() == "macos")
                {
                    var command = "open " + newFile;
                    Utility.ExecuteCommand(command);
                }
                else
                {
                    Process.Start(newFile.ToString());
                }
            }
            else
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }

        private static void ExportClassifications(LiteCollection<KeyRatio> dbKeyRatio)
        {
            Console.Write("Exporting...");
            var all = dbKeyRatio.Find(Query.All());
            var newFile = new FileInfo(GetFileName());
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Data");
                worksheet.Cells["A1"].Value = "Ticker";
                worksheet.Cells["B1"].Value = "Company Name";

                worksheet.Cells["C1"].Value = "Sector";
                worksheet.Cells["D1"].Value = "Industry";
                worksheet.Cells["E1"].Value = "Subindustry";

                //Format cells
                worksheet.Cells[1, 1, 1, 5].Style.Font.Bold = true;
                
                //Freeze both header pane and ticker column
                if (Settings.FreezeHeaderPane && Settings.FreezeTickerColumn)
                {
                    worksheet.View.FreezePanes(2, 2);
                }
                //Freeze just the header pane
                else if (Settings.FreezeHeaderPane && !Settings.FreezeTickerColumn)
                {
                    worksheet.View.FreezePanes(2, 1);
                }
                //Freeze just the ticker column
                else if (!Settings.FreezeHeaderPane && Settings.FreezeTickerColumn)
                {
                    worksheet.View.FreezePanes(1, 2);
                }

                var row = 2;
                foreach (var a in all)
                {
                    worksheet.Cells[row, 1].Value = a.Ticker;
                    worksheet.Cells[row, 2].Value = a.Name;
                    worksheet.Cells[row, 3].Value = a.Sector;
                    worksheet.Cells[row, 4].Value = a.Industry;
                    worksheet.Cells[row, 5].Value = a.Subindustry;
                    row++;
                }

                //Auto fit all columns
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                try
                {
                    package.SaveAs(newFile);
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                    Console.WriteLine("\rSomething went wrong...");
                    Console.WriteLine("Check the error log for more details");
                    Console.ReadKey();
                    return;
                }
            }

            Console.WriteLine("\rExcel file generated successfully!");
            if (Settings.AutoLaunchExcel)
            {
                Console.WriteLine("Launching Excel...");
                if (Utility.DetermineOs() == "macos")
                {
                    var command = "open " + newFile;
                    Utility.ExecuteCommand(command);
                }   
                else
                {
                    Process.Start(newFile.ToString());
                }
            }
            else
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }

        #endregion

        #region SearchHelperMethods

        private static int IndexAt(string[] array, string key)
        {
            for (var i = 0; i < array.Length; i++)
            {
                if (array[i] == key)
                {
                    return i;
                }
            }
            return -1;
        }

        private static int ContainsEntry(string[] container, string key)
        {
            for (var i = 0; i < container.Length; i++)
            {
                if (GetYear(container[i]) == int.Parse(key))
                    return i;
            }
            return int.MinValue;
        }

        private static int GetYear(string date)
        {
            try
            {
                var array = date.Split('-');
                return int.Parse(array[0]);
            }
            catch (Exception)
            {
                return -1;
            }
        }

        private static void FormatResults(List<string> array)
        {
            Console.WriteLine(string.Join(", ", array.Take(array.Count - 1)) + (array.Count <= 1 ? "" : " and ") + array.LastOrDefault());
        }

        private static List<Dictionary<string, int>> SortYears(IReadOnlyList<KeyRatio> results, string[] years)
        {
            //Inflate
            var pos = new List<Dictionary<string, int>>();
            for (var i = 0; i < results.Count; i++)
            {
                var y = years.ToDictionary(t => t, t => int.MinValue);
                pos.Add(y);
            }

            //Populate
            for (var i = 0; i < results.Count; i++)
            {
                var curr = results[i];
                foreach (var t in years)
                {
                    if (ContainsEntry(curr.FiscalYears, t) != int.MinValue)
                    {
                        pos[i][t] = ContainsEntry(curr.FiscalYears, t);
                    }
                }
            }
            return pos;
        }

        #endregion

        #region ExcelHelperMethods

        private static void WriteToExcel(List<KeyRatio> results, List<string> metrics, string[] years, IReadOnlyList<Dictionary<string, int>> pos, bool freeze)
        {
            var newFile = new FileInfo(GetFileName());

            AppendToExcel(newFile, true, results, metrics, years, pos, 1, freeze);

            Console.WriteLine("File generated successfully!");
            if (Settings.AutoLaunchExcel)
            {
                Console.WriteLine("Launching Excel...");
                if (Utility.DetermineOs() == "macos")
                {
                    var command = "open " + newFile;
                    Utility.ExecuteCommand(command);
                }
                else
                {
                    Process.Start(newFile.ToString());
                }
            }
            else
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }

        private static int AppendToExcel(FileInfo file, bool createNewSheet, List<KeyRatio> results, List<string> metrics, string[] years, IReadOnlyList<Dictionary<string, int>> pos, int header, bool freeze)
        {
            int row;
            using (var package = new ExcelPackage(file))
            {
                var worksheet = createNewSheet ? package.Workbook.Worksheets.Add("Data") : package.Workbook.Worksheets[1];

                var col = 1;
                WriteCell(ref worksheet, header, ref col, "Ticker");
                WriteCell(ref worksheet, header, ref col, "Company Name");

                if (Settings.IncludeCompanyInfo)
                {
                    WriteCell(ref worksheet, header, ref col, "Sector");
                    WriteCell(ref worksheet, header, ref col, "Industry");
                    WriteCell(ref worksheet, header, ref col, "Subindustry");
                    
                }
                WriteCell(ref worksheet, header, ref col, "Metric");
                WriteCell(ref worksheet, header, ref col, "Currency");
                WriteCell(ref worksheet, header, ref col, "Fiscal Year Ends");

                foreach (var year in years)
                {
                    var tmpCol = col;
                    FormatCell(ref worksheet, header, tmpCol, "number");
                    WriteCell(ref worksheet, header, ref col, year);
                    tmpCol++;
                }

                //Format cells
                worksheet.Cells[header, 1, header, col].Style.Font.Bold = true;
                if (freeze)
                {
                    //Freeze both header pane and ticker column
                    if (Settings.FreezeHeaderPane && Settings.FreezeTickerColumn)
                    {
                        worksheet.View.FreezePanes(2, 2);
                    }
                    //Freeze just the header pane
                    else if (Settings.FreezeHeaderPane && !Settings.FreezeTickerColumn)
                    {
                        worksheet.View.FreezePanes(2, 1);
                    }
                    //Freeze just the ticker column
                    else if (!Settings.FreezeHeaderPane && Settings.FreezeTickerColumn)
                    {
                        worksheet.View.FreezePanes(1, 2);
                    }
                }

                row = header + 1;
                col = 1;
                var currentCompany = 0;
                var formatDictionary = 
                    File.ReadAllLines(Data + "FormatKey.txt")
                    .Select(s => s.Split(','))
                    .ToDictionary(split => split[0], split => split[1]);

                foreach (var keyRatio in results)
                {
                    foreach (var metric in metrics)
                    {
                        WriteCell(ref worksheet, row, ref col, keyRatio.Ticker);
                        WriteCell(ref worksheet, row, ref col, keyRatio.Name);
                        if (Settings.IncludeCompanyInfo)
                        {
                            WriteCell(ref worksheet, row, ref col, keyRatio.Sector);
                            WriteCell(ref worksheet, row, ref col, keyRatio.Industry);
                            WriteCell(ref worksheet, row, ref col, keyRatio.Subindustry);
                        }
                        //Metrics
                        var position = IndexAt(_key, metric);
                        var tmp = _value[position];
                        WriteCell(ref worksheet, row, ref col, tmp);
                        WriteCell(ref worksheet, row, ref col, keyRatio.CurrencyType);
                        WriteCell(ref worksheet, row, ref col, keyRatio.FiscalYearEnd);

                        //Years
                        var metricArray = keyRatio.FinancialData[metric];
                        
                        var type = formatDictionary[metric];
                        foreach (var t in years)
                        {
                            if (pos[currentCompany][t] == int.MinValue)
                            {
                                WriteCell(ref worksheet, row, ref col, "N/A");
                            }
                            else
                            {
                                FormatCell(ref worksheet, row, col, type);
                                double number;
                                try
                                {
                                    number = metricArray[pos[currentCompany][t]];
                                }
                                catch (Exception e)
                                {
                                    Utility.WriteToErrorLog(e);
                                    number = 0.00;
                                }
                                if (type == "percent")
                                {
                                    number = number * 0.01;
                                }
                                WriteCell(ref worksheet, row, ref col, number);
                            }
                        }
                        col = 1;
                        row++;
                    }
                    currentCompany++;
                }

                //Auto fit all columns
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                try
                {
                    package.SaveAs(file);
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                }
            }
            return row + Settings.MultipleSearchSpacing;
        }

        private static void WriteCell(ref ExcelWorksheet worksheet, int row, ref int column, string text)
        {
            worksheet.Cells[row, column].Value = text;
            column++;
        }

        private static void WriteCell(ref ExcelWorksheet worksheet, int row, ref int column, double num)
        {
            worksheet.Cells[row, column].Value = num;
            column++;
        }

        /// <summary>
        /// Formats cell according to the type of data stored in it
        /// </summary>
        /// <param name="worksheet">Worksheet object</param>
        /// <param name="row">Row of the cell to format</param>
        /// <param name="column">Column of the cell to format</param>
        /// <param name="type">Type of format; valid types: "number", "decimal", "percent"</param>
        private static void FormatCell(ref ExcelWorksheet worksheet, int row, int column, string type)
        {
            switch (type)
            {
                case "number":
                    worksheet.Cells[row, column].Style.Numberformat.Format = "#";
                    break;
                case "decimal":
                    //2 decimal places defaults
                    worksheet.Cells[row, column].Style.Numberformat.Format = "#,##0.00";
                    break;
                case "percent":
                    //Percentage (1 = 100%, 0.01 = 1%)
                    worksheet.Cells[row, column].Style.Numberformat.Format = "0.00%";
                    break;
            }

            /* 
            CELL FORMATTING:
            //integer (not really needed unless you need to round numbers, Excel with use default cell properties)
            ws.Cells["A1:A25"].Style.Numberformat.Format = "0";

            //integer without displaying the number 0 in the cell
            ws.Cells["A1:A25"].Style.Numberformat.Format = "#";

            //number with 1 decimal place
            ws.Cells["A1:A25"].Style.Numberformat.Format = "0.0";

            //number with 2 decimal places
            ws.Cells["A1:A25"].Style.Numberformat.Format = "0.00";

            //number with 2 decimal places and thousand separator
            ws.Cells["A1:A25"].Style.Numberformat.Format = "#,##0.00";

            //number with 2 decimal places and thousand separator and money symbol
            ws.Cells["A1:A25"].Style.Numberformat.Format = "€#,##0.00";

            //percentage (1 = 100%, 0.01 = 1%)
            ws.Cells["A1:A25"].Style.Numberformat.Format = "0%";

            //default DateTime pattern
            worksheet.Cells["A1:A25"].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

            //custom DateTime pattern
            worksheet.Cells["A1:A25"].Style.Numberformat.Format = "dd-MM-yyyy HH:mm";
            */
        }

        #endregion

        private static string GetFileName()
        {
            if (!File.Exists(Desktop + "Query Results.xlsx")) return Desktop + "Query Results.xlsx";
            var fileSuffix = 1;
            while (true)
            {
                if (!File.Exists(Desktop + "Query Results" + " (" + fileSuffix + ").xlsx"))
                {
                    return Desktop + "Query Results" + " (" + fileSuffix + ").xlsx";
                }
                fileSuffix++;
            }
        }
    }
}
