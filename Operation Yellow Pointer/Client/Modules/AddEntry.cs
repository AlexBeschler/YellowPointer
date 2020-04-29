using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Timers;
using HtmlAgilityPack;
using LiteDB;
using OfficeOpenXml;
using Operation_Yellow_Pointer.Client.Platform;
using Timer = System.Timers.Timer;

namespace Operation_Yellow_Pointer.Client.Modules
{
    internal static class AddEntry
    {
        internal static string Database;
        internal static string Desktop;
        internal static string Data;
        internal static string Temp;

        private static int _line;
        private static int _total;
        private static int _successfulAdd;

        private static int _currentYear;

        private const string ErrorLog = "AddToDatabaseErrorList.txt";

        public static void Init()
        {
            _successfulAdd = 0;
            _currentYear = int.Parse(Convert
                .ToDateTime(DateTime.Now.ToString(CultureInfo.CurrentCulture))
                .Year
                .ToString());
        }

        public static void Menu()
        {
            _successfulAdd = 0;
            Console.Title = "Add Database Entries";
            do
            {
                Console.Clear();
                Console.WriteLine("Add Database Entries:");
                Console.WriteLine("1. Import batch additions from Excel");
                Console.WriteLine("2. Add specific entry");
                Console.WriteLine("3. Back");
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
                        return;
                    default:
                        Console.WriteLine("Invalid input");
                        break;
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
                var dataFolder = CreateFolder(Desktop, "DownloadedData");
                
                for (; _line <= _total; _line++)
                {
                    if (!string.IsNullOrEmpty(worksheet.Cells[_line, 1].Value.ToString()))
                        Add(worksheet.Cells[_line, 1].Value.ToString().Trim(), true, dataFolder);
                }
                timer.Enabled = false;
                Console.WriteLine();
                Console.WriteLine("Success!");
                Console.WriteLine("Added " + _successfulAdd + " new entries.");
                Console.WriteLine("Press any key to continue");
                Console.ReadKey();
            }
        }

        private static void Manual()
        {
            string ticker;
            var success = false;
            do
            {
                Console.WriteLine("Ticker: ");
                ticker = Console.ReadLine();
                switch (ticker)
                {
                    case null:
                        Console.WriteLine("Invalid ticker");
                        break;
                    case "back":
                        return;
                    default:
                        ticker = ticker.ToUpper();
                        success = true;
                        break;
                }
            } while (!success);
            success = Add(ticker, false, null);
            if (success)
            {
                Console.WriteLine("Success! Press any key to continue");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }

        private static bool Add(string ticker, bool isBatchJob, string dataFolder)
        {
            using (var db = new LiteDatabase(Database))
            {
                var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                //Check to see if duplicate company
                var result = dbKeyRatio.FindOne(Query.EQ("Ticker", ticker));
                if (result != null)
                {
                    Utility.WriteToLog(ticker, "Ticker already exists in database", ErrorLog);
                    return false;
                }
                if (!isBatchJob)
                {
                    Console.WriteLine("Downloading files...");
                }
                if (dataFolder == null)
                    dataFolder = Desktop;
                if (!Download(ticker, dataFolder))
                {
                    if (!isBatchJob)
                        Console.WriteLine("Failed to download Morningstar data");
                    return false;
                }

                if (!isBatchJob)
                {
                    Console.WriteLine("Extracting Data...");
                }
                var keyRatio = ParseExcel(ticker, dataFolder);
                
                if (keyRatio == null)
                {
                    if (isBatchJob) return false;
                    Console.WriteLine("Program failed to extract data");
                    Console.WriteLine("See error log on desktop for more details");
                    return false;
                }
                var cik = GetCik(ticker);
                if (cik == null)
                {
                    if (!isBatchJob)
                        Console.WriteLine("Failed to get CIK");
                }
                else
                {
                    keyRatio.Cik = cik;
                }
                if (!isBatchJob)
                {
                    Console.WriteLine("Adding to database...");
                }
                try
                {
                    dbKeyRatio.Insert(keyRatio);
                    _successfulAdd++;
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                }
            }
            return true;
        }

        public static bool Add(string ticker, bool amend)
        {
            using (var db = new LiteDatabase(Database))
            {
                var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");

                var result = dbKeyRatio.FindOne(Query.EQ("Ticker", ticker));
                if (result == null)
                {
                    Utility.WriteToLog(ticker, "Could not find ticker when updating financial information", ErrorLog);
                    return false;
                }

                if (!Download(ticker))
                    return false;

                var success = ParseExcel(ticker, ref result, amend);

                if (!success) return false;
                try
                {
                    dbKeyRatio.Update(result);
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                }
            }
            return true;
        }

        private static bool Download(string ticker, string dataFolder)
        {
            const string kr = "http://financials.morningstar.com/ajax/exportKR2CSV.html?t=";
            const string report = "http://financials.morningstar.com/ajax/ReportProcess4CSV.html?t=";
            const string parameters1 = "&reportType=";
            const string parameters2 = "&period=12&dataType=A&order=asc&columnYear=10&number=3";
            const int maxTimeout = 3;
            var timeout = 0;

            //No need to download if data is already there
            if (File.Exists(dataFolder + ticker.Trim() + ".xlsx")) return true;
            
            using (var download = new Utility.WebDownload())
            {
                bool success;
                do
                {
                    try
                    {
                        //DL KeyRatios
                        download.DownloadFile(
                            kr +
                            ticker,
                            Temp +
                            "KeyRatios.csv");
                        //DL Income Statement
                        download.DownloadFile(
                            report +
                            ticker +
                            parameters1 +
                            "is" +
                            parameters2,
                            Temp +
                            "IncomeStatement.csv");
                        //DL Balance Sheet
                        download.DownloadFile(
                            report +
                            ticker +
                            parameters1 +
                            "bs" +
                            parameters2,
                            Temp +
                            "BalanceSheet.csv");
                        //DL Cash Flows
                        download.DownloadFile(
                            report +
                            ticker +
                            parameters1 +
                            "cf" +
                            parameters2,
                            Temp +
                            "CashFlows.csv");
                        success = true;
                    }
                    catch (Exception e)
                    {
                        Utility.WriteToErrorLog(e);
                        success = false;
                        if (timeout == maxTimeout)
                        {
                            return false;
                        }
                        var time = timeout + 1;
                        Console.Write("\rDownload failed, trying " + time + " of 3");
                        Thread.Sleep(1000); //Sleep for 1 second
                        timeout++;
                    }
                } while (!success);
            }
            //Combine to one XLSX
            if (dataFolder == null)
            {
                dataFolder = Desktop;
            }
            using (var package = new ExcelPackage(new FileInfo(dataFolder + ticker.Trim() + ".xlsx")))
            {
                //Add Key Ratios
                var source = Temp + "KeyRatios.csv";
                ExcelWorksheet worksheet;
                try
                {
                    worksheet = package.Workbook.Worksheets.Add("Key Ratios");
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                    worksheet = package.Workbook.Worksheets[1];
                }
                var data = GetCsvData(source);
                if (data.Count < 4)
                {
                    Utility.WriteToLog(ticker, "Blank or unstandardized Excel document", ErrorLog);
                    return false;
                }
                worksheet.Cells[1, 1].LoadFromArrays(data);

                //Add income statement
                source = Temp + "IncomeStatement.csv";
                try
                {
                    worksheet = package.Workbook.Worksheets.Add("Income Statement");
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                    worksheet = package.Workbook.Worksheets[2];
                }
                data = GetCsvData(source);
                if (data.Count < 4)
                {
                    Utility.WriteToLog(ticker, "Blank or unstandardized Excel document", ErrorLog);
                    return false;
                }
                worksheet.Cells[1, 1].LoadFromArrays(data);

                //Add balance sheet
                source = Temp + "BalanceSheet.csv";
                try
                {
                    worksheet = package.Workbook.Worksheets.Add("Balance Sheet");
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                    worksheet = package.Workbook.Worksheets[3];
                }
                data = GetCsvData(source);
                if (data.Count < 4)
                {
                    Utility.WriteToLog(ticker, "Blank or unstandardized Excel document", ErrorLog);
                    return false;
                }
                worksheet.Cells[1, 1].LoadFromArrays(data);

                //Add cash flows
                source = Temp + "CashFlows.csv";
                try
                {
                    worksheet = package.Workbook.Worksheets.Add("Cash Flows");
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                    worksheet = package.Workbook.Worksheets[4];
                }
                data = GetCsvData(source);
                if (data.Count < 4)
                {
                    Utility.WriteToLog(ticker, "Blank or unstandardized Excel document", ErrorLog);
                    return false;
                }
                worksheet.Cells[1, 1].LoadFromArrays(data);

                try
                {
                    package.Save();
                    File.Delete(Temp + "KeyRatios.csv");
                    File.Delete(Temp + "IncomeStatement.csv");
                    File.Delete(Temp + "BalanceSheet.csv");
                    File.Delete(Temp + "CashFlows.csv");
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                }
            }
            return true;
        }

        private static bool Download(string ticker)
        {
            const string kr = "http://financials.morningstar.com/ajax/exportKR2CSV.html?t=";
            const int maxTimeout = 3;
            var timeout = 0;
            
            using (var download = new Utility.WebDownload())
            {
                bool success;
                do
                {
                    try
                    {
                        //DL KeyRatios
                        download.DownloadFile(kr + ticker, Temp + "KeyRatios.csv");
                        success = true;
                    }
                    catch (Exception e)
                    {
                        Utility.WriteToErrorLog(e);
                        success = false;
                        if (timeout == maxTimeout)
                        {
                            return false;
                        }
                        var time = timeout + 1;
                        Console.Write("\rDownload failed, trying " + time + " of 3");
                        Thread.Sleep(1000); //Sleep for 1 second
                        timeout++;
                    }
                } while (!success);
            }
            //Combine to one XLSX
            using (var package = new ExcelPackage(new FileInfo(Temp + ticker.Trim() + ".xlsx")))
            {
                //Add Key Ratios
                var source = Temp + "KeyRatios.csv";
                ExcelWorksheet worksheet;
                try
                {
                    worksheet = package.Workbook.Worksheets.Add("Key Ratios");
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                    worksheet = package.Workbook.Worksheets[1];
                }
                var data = GetCsvData(source);
                if (data.Count < 4)
                {
                    Utility.WriteToLog(ticker, "Blank or unstandardized Excel document", ErrorLog);
                    return false;
                }
                worksheet.Cells[1, 1].LoadFromArrays(data);

                try
                {
                    package.Save();
                    File.Delete(Temp + "KeyRatios.csv");
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                }
            }
            return true;
        }

        private static KeyRatio ParseExcel(string ticker, string dataFolder)
        {
            var keyRatio = new KeyRatio();

            //Open Excel File
            var file = new FileInfo(dataFolder + ticker + ".xlsx");
            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets[1];
                //Get the name of company
                string name;
                try
                {
                    name = worksheet.Cells[1, 1].Value.ToString().Trim();
                    name = name.Substring(46);
                }
                catch (NullReferenceException)
                {
                    Utility.WriteToLog(ticker, "Blank Excel document", ErrorLog);
                    return null;
                }
                catch (ArgumentOutOfRangeException) //Catches "N/A" from QuantizeString method
                {
                    Utility.WriteToLog(ticker, "Blank Excel document", ErrorLog);
                    return null;
                }
                catch (Exception)
                {
                    Utility.WriteToLog(ticker, "Blank or unstandardized Excel document", ErrorLog);
                    return null;
                }
                keyRatio.Name = name;
                keyRatio.Ticker = ticker;

                //Currency type
                string currency;
                try
                {
                    currency = worksheet.Cells[4, 1].Value.ToString().Trim();
                    if (!currency.Contains("Revenue")) //Data is inconsistent
                    {
                        Utility.WriteToLog(ticker, "Data is inconsistent", ErrorLog);
                        return null;
                    }
                    currency = currency.Substring(8, 3); //Grabs currency type
                }
                catch (ArgumentOutOfRangeException)
                {
                    Utility.WriteToLog(ticker, "Morningstar didn't have data on selected ticker", ErrorLog);
                    return null;
                }
                catch (Exception)
                {
                    Utility.WriteToLog(ticker, "Morningstar didn't have data on selected ticker", ErrorLog);
                    return null;
                }
                keyRatio.CurrencyType = currency;

                //Fiscal Years
                //Gets the first date on the spreadsheet, gets the month from the date,
                //then converts the month number to a word; basically: 2007-12  --> December
                keyRatio.FiscalYearEnd =
                    GetMonthName(
                        GetMonth(
                            worksheet.Cells[3, 2].Value.ToString()).ToString());

                keyRatio.FiscalYears = new string[10];
                for (var i = 2; i < 12; i++)
                {
                    keyRatio.FiscalYears[i - 2] = worksheet.Cells[3, i].Value.ToString();
                }

                //Inflate dictionaries
                var locNormal = File.ReadAllLines(Data + "LOC_normal.txt");
                var locDynamic = File.ReadAllLines(Data + "LOC_dynamic.txt");

                var normal = locNormal.Select(v => v.Split(',')).ToDictionary(tmp => tmp[1], tmp => tmp[0]);
                var dynam = locDynamic.Select(v => v.Split(',')).ToDictionary(tmp => tmp[1], tmp => tmp[0] + " " + currency);

                var totalRow = worksheet.Dimension.End.Row;
                keyRatio.FinancialData = new Dictionary<string, double[]>();

                //This number is for when finding the Key Ratio -> Growth rows in the spreadsheet
                var growthRatioRow = 0;

                //Find all metrics that cannot be found in spreadsheet
                var wsMetrics = new List<string>();
                for (var i = 1; i <= totalRow; i++)
                {
                    var header = worksheet.Cells[i, 1].Value;
                    if (header == null)
                    {
                        continue;
                    }
                    if (header.ToString() == "Revenue %")
                        growthRatioRow = i + 1; //References first meaningful line
                    wsMetrics.Add(header.ToString());
                }

                foreach (var s in normal)
                {
                    if (!wsMetrics.Any(s.Value.Contains))
                    {
                        Utility.WriteToLog(ticker, s + " not found", ErrorLog);
                    }
                }
                foreach (var d in dynam)
                {
                    if (!wsMetrics.Exists(x => x.Contains(d.Value)))
                    {
                        Utility.WriteToLog(ticker, d + " not found", ErrorLog);
                    }
                }

                //Loop to get all normal metrics
                //Uses "Matches" method
                for (var i = 1; i <= totalRow; i++)
                {
                    var header = worksheet.Cells[i, 1].Value;
                    if (header == null)
                        continue;
                    if (!normal.ContainsValue(header.ToString()))
                        continue;
                    keyRatio.FinancialData.Add(
                        normal.FirstOrDefault(x => x.Value.Equals(header.ToString())).Key, GetData(worksheet, i));
                }
                //Loop to get all metrics with currency in name
                for (var i = 1; i <= totalRow; i++)
                {
                    var header = worksheet.Cells[i, 1].Value;
                    if (header == null)
                        continue;
                    if (!dynam.Any(x => header.ToString().Contains(x.Value)))
                        continue;
                    keyRatio.FinancialData.Add(
                        dynam.FirstOrDefault(x => header.ToString().Contains(x.Value)).Key, GetData(worksheet, i));
                }

                //Revenue %
                keyRatio.FinancialData.Add("RGYY", GetData(worksheet, growthRatioRow));
                keyRatio.FinancialData.Add("RGY3", GetData(worksheet, growthRatioRow + 1));
                keyRatio.FinancialData.Add("RGY5", GetData(worksheet, growthRatioRow + 2));
                keyRatio.FinancialData.Add("RGY10", GetData(worksheet, growthRatioRow + 3));

                //Operating Income %
                keyRatio.FinancialData.Add("OIGYY", GetData(worksheet, growthRatioRow + 5));
                keyRatio.FinancialData.Add("OIGY3", GetData(worksheet, growthRatioRow + 6));
                keyRatio.FinancialData.Add("OIGY5", GetData(worksheet, growthRatioRow + 7));
                keyRatio.FinancialData.Add("OIGY10", GetData(worksheet, growthRatioRow + 8));

                //Net Income %
                keyRatio.FinancialData.Add("NIGYY", GetData(worksheet, growthRatioRow + 10));
                keyRatio.FinancialData.Add("NIGY3", GetData(worksheet, growthRatioRow + 11));
                keyRatio.FinancialData.Add("NIGY5", GetData(worksheet, growthRatioRow + 12));
                keyRatio.FinancialData.Add("NIGY10", GetData(worksheet, growthRatioRow + 13));
                
                
                //EPS %
                keyRatio.FinancialData.Add("EPSGYY", GetData(worksheet, growthRatioRow + 15));
                keyRatio.FinancialData.Add("EPSGY3", GetData(worksheet, growthRatioRow + 16));
                keyRatio.FinancialData.Add("EPSGY5", GetData(worksheet, growthRatioRow + 17));
                keyRatio.FinancialData.Add("EPSGY10", GetData(worksheet, growthRatioRow + 18));
            }
            return keyRatio;
        }

        private static bool ParseExcel(string ticker, ref KeyRatio result, bool amend)
        {
            //Open Excel File
            var file = new FileInfo(Temp + ticker + ".xlsx");
            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets[1];
                //Checks for inconsistent spreadsheet
                try
                {
                    var name = worksheet.Cells[1, 1].Value.ToString().Trim();
                    name = name.Substring(46);
                }
                catch (NullReferenceException)
                {
                    Utility.WriteToLog(ticker, "Blank Excel document", ErrorLog);
                    return false;
                }
                catch (ArgumentOutOfRangeException) //Catches "N/A" from QuantizeString method
                {
                    Utility.WriteToLog(ticker, "Blank Excel document", ErrorLog);
                    return false;
                }
                catch (Exception)
                {
                    Utility.WriteToLog(ticker, "Blank or unstandardized Excel document", ErrorLog);
                    return false;
                }
                string currency;
                try
                {
                    currency = worksheet.Cells[4, 1].Value.ToString().Trim();
                    if (!currency.Contains("Revenue")) //Data is inconsistent
                    {
                        Utility.WriteToLog(ticker, "Data is inconsistent", ErrorLog);
                        return false;
                    }
                    currency = currency.Substring(8, 3); //Grabs currency type
                }
                catch (ArgumentOutOfRangeException)
                {
                    Utility.WriteToLog(ticker, "Morningstar didn't have data on selected ticker", ErrorLog);
                    return false;
                }
                catch (Exception)
                {
                    Utility.WriteToLog(ticker, "Morningstar didn't have data on selected ticker", ErrorLog);
                    return false;
                }

                //Check if Morningstar has current financial information
                //This has been an error in the past:
                //  The SEC list had the 10-K filed with it, but Morningstar didn't have the updated data yet
                try
                {
                    var yearDate = GetYear(worksheet.Cells[3, 11].Value.ToString());
                    if (yearDate < _currentYear && !amend)
                    {
                        Console.WriteLine("Skipping " + ticker + "...");
                        Utility.WriteToLog(ticker, "Skipping, as Morningstar didn't have up-to-date data on it", ErrorLog);
                        return false;
                    }
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                }

                //Load fiscal years to new array
                if (!amend)
                {
                    var fYears = new string[result.FiscalYears.Length + 1];
                    for (var i = 0; i < result.FiscalYears.Length; i++)
                    {
                        fYears[i] = result.FiscalYears[i];
                    }
                    var refIndex = fYears.Length - 1;
                    fYears[refIndex] = worksheet.Cells[3, 11].Value.ToString();
                    result.FiscalYears = fYears;
                }

                //Inflate dictionaries
                var locNormal = File.ReadAllLines(Data + "LOC_normal.txt");
                var locDynamic = File.ReadAllLines(Data + "LOC_dynamic.txt");

                var normal = locNormal.Select(v => v.Split(',')).ToDictionary(tmp => tmp[1], tmp => tmp[0]);
                var dynam = locDynamic.Select(v => v.Split(',')).ToDictionary(tmp => tmp[1], tmp => tmp[0] + " " + currency);

                var totalRow = worksheet.Dimension.End.Row;

                //This number is for when finding the Key Ratio -> Growth rows in the spreadsheet
                var growthRatioRow = 0;

                //Find all metrics that cannot be found in spreadsheet
                var wsMetrics = new List<string>();
                for (var i = 1; i <= totalRow; i++)
                {
                    var header = worksheet.Cells[i, 1].Value;
                    if (header == null)
                    {
                        continue;
                    }
                    if (header.ToString() == "Revenue %")
                        growthRatioRow = i + 1; //References first meaningful line
                    wsMetrics.Add(header.ToString());
                }

                foreach (var s in normal)
                {
                    if (!wsMetrics.Any(s.Value.Contains))
                    {
                        Utility.WriteToLog(ticker, s + " not found", ErrorLog);
                    }
                }
                foreach (var d in dynam)
                {
                    if (!wsMetrics.Exists(x => x.Contains(d.Value)))
                    {
                        Utility.WriteToLog(ticker, d + " not found", ErrorLog);
                    }
                }

                //Loop to get all normal metrics
                for (var i = 1; i <= totalRow; i++)
                {
                    var header = worksheet.Cells[i, 1].Value;
                    if (header == null)
                        continue;
                    if (!normal.ContainsValue(header.ToString()))
                        continue;
                    var previousDataKey = normal.FirstOrDefault(x => x.Value.Equals(header.ToString())).Key;
                    result.FinancialData[previousDataKey] = 
                        GetData(worksheet, result.FinancialData[previousDataKey], i, amend);
                }
                //Loop to get all metrics with currency in name
                for (var i = 1; i <= totalRow; i++)
                {
                    var header = worksheet.Cells[i, 1].Value;
                    if (header == null)
                        continue;
                    if (!dynam.Any(x => header.ToString().Contains(x.Value)))
                        continue;
                    var previousDataKey = dynam.FirstOrDefault(x => header.ToString().Contains(x.Value)).Key;
                    result.FinancialData[previousDataKey] =
                        GetData(worksheet, result.FinancialData[previousDataKey], i,amend);
                }

                var percentArray = new[]
                {
                    "RGYY",
                    "RGY3",
                    "RGY5",
                    "RGY10",
                    "OIGYY",
                    "OIGY3",
                    "OIGY5",
                    "OIGY10",
                    "NIGYY",
                    "NIGY3",
                    "NIGY5",
                    "NIGY10",
                    "EPSGYY",
                    "EPSGY3",
                    "EPSGY5",
                    "EPSGY10"
                };
                var pos = 0;
                //16 stands for how many rows we need
                for (var i = 0; i < 19; i++)
                {
                    var index = percentArray[pos];
                    if (i == 4 || i == 9 || i == 14)
                    {
                        i++;
                    }
                    result.FinancialData[index] = GetData(worksheet, result.FinancialData[index], growthRatioRow + i, amend);
                    pos++;
                }
            }
            return true;
        }

        private static double[] GetData(ExcelWorksheet worksheet, int row)
        {
            var list = new double[10];
            for (var k = 2; k < 12; k++)
            {
                var data = worksheet.Cells[row, k].Value;
                if (data == null)
                {
                    list[k - 2] = 0.0;
                    continue;
                }
                if (double.TryParse(data.ToString(), out var result))
                    list[k - 2] = result;
                else
                    list[k - 2] = 0.0;
            }
            return list;
        }

        private static double[] GetData(ExcelWorksheet worksheet, double[] previousData, int row, bool amend)
        {
            int lastIndex;
            double[] list;
            if (amend)
            {
                list = new double[previousData.Length];
                lastIndex = previousData.Length - 1;
            }
            else
            {
                list = new double[previousData.Length + 1];
                lastIndex = previousData.Length;
            }
            
            //Inflate the new array
            for (var i = 0; i < previousData.Length; i++)
            {
                list[i] = previousData[i];
            }

            var data = worksheet.Cells[row, 11].Value;
            if (data == null)
            {
                list[lastIndex] = 0.0;
            }
            else
            {
                if (double.TryParse(data.ToString(), out var result))
                    list[lastIndex] = result;
                else
                    list[lastIndex] = 0.0;
            }
            return list;
        }

        private static void UpdateProgress(object source, ElapsedEventArgs e)
        {
            //(Actual Work / Work) * 100
            var progress = _line / (double)_total;
            progress *= 100.00;
            Console.Write("\r                ");
            Console.Write("\rProgress: " + Math.Round(progress, 2) + "%");
        }

        private static int GetMonth(string date)
        {
            try
            {
                var array = date.Split('-');
                return int.Parse(array[1]);
            }
            catch (Exception)
            {
                return -1;
            }
        }
        private static string GetMonthName(string date)
        {
            int.TryParse(date, out var month);
            return CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
        }

        private static List<object[]> GetCsvData(string file)
        {
            var lines = File.ReadAllLines(file);
            var array = new List<object[]>();

            foreach (var line in lines)
            {
                var lineArray = line.ToCharArray();
                var i = 0;
                var record = new List<char>();
                var obj = new List<object>();

                while (i < lineArray.Length)
                {
                    switch (lineArray[i])
                    {
                        case ',':
                            obj.Add(GetString(record)); //Convert to object array
                            record.Clear();
                            i++;
                            continue; //Continue, there's more records
                        case '"':
                            i++; //So it doesn't add bracket
                            var foundEos = false;
                            while (!foundEos)
                            {
                                if (lineArray[i] == '"')
                                {
                                    i++; //Again, to prevent bracket adding
                                    foundEos = true;
                                }
                                else
                                {
                                    record.Add(lineArray[i]);
                                    i++;
                                }
                            }
                            continue;
                    }
                    record.Add(lineArray[i]);
                    i++;
                }
                obj.Add(GetString(record)); //To perserve last record
                array.Add(obj.ToArray()); //Add row
            }
            return array;
        }

        private static string GetString(List<char> list)
        {
            var sb = new StringBuilder();
            foreach (var c in list)
            {
                sb.Append(c);
            }
            return sb.ToString();
        }

        private static string CreateFolder(string parentDirectory, string folderName)
        {
            Directory.CreateDirectory(parentDirectory + "DownloadedData");
            var separator = Utility.DetermineOs() == "windows" ? '\\' : '/';
            return parentDirectory + folderName + separator;
        }

        private static string GetCik(string ticker)
        {
            const string cikLink1 = "https://www.sec.gov/cgi-bin/browse-edgar?CIK=";
            const string cikLink2 = "&owner=exclude&action=getcompany";
            var html = cikLink1 + ticker + cikLink2;
            try
            {
                var web = new HtmlWeb();
                var htmlDoc = web.Load(html);
                var result = htmlDoc.DocumentNode.SelectSingleNode("//div[4]/div[1]/div[3]/span[1]/a[1]").InnerHtml;
                return ExtractCik(result);
            }
            catch (Exception e)
            {
                Utility.WriteToErrorLog(e);
                Utility.WriteToLog(ticker, "Could not find CIK. This ticker will NOT update automatically.", ErrorLog);
            }
            return null;
        }

        private static string ExtractCik(string hrefText)
        {
            return Regex.Match(hrefText, @"\d+").Value;
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
    }
}
