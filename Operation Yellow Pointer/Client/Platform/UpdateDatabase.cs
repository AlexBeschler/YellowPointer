using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using LiteDB;
using Operation_Yellow_Pointer.Client.Modules;

namespace Operation_Yellow_Pointer.Client.Platform
{
    internal class UpdateDatabase
    {
        internal static string Temp { private get; set; }
        internal static string Database { private get; set; }

        private const string SecFormAddress = "https://www.sec.gov/Archives/edgar/full-index/form.gz";
        private static string ErrorLog = "UpdateFinancialInfoError.txt";

        //Types of forms that the program scans for
        private static readonly string[] Forms = { "10-K" };
        private static readonly string[] AmendForms = {"10-K/A"};

        public static void Update()
        {
            if (!DownloadFormsIndex()) return;
            var filings = ParseFormsIndex();
            var amendments = Get10KAmendments();

            if (filings.Count == 0) //If there's no filings relevant to updating the database
            {
                Console.WriteLine("No forms have been filed at this time");
                return;
            }
            var tickersToUpdate = EndingFiscalYearCompanies();
            
            //Queue of companies to DL from Morningstar
            //var queue = tickersToUpdate.Where(x => filings.Contains(x)).ToList();
            var queue = filings.Where(filing => tickersToUpdate.Contains(filing)).ToList();

            Console.Write("There ");
            var count = queue.Count + amendments.Count;
            Console.Write(count == 1 ? "is " : "are ");
            Console.Write(count + " ticker");
            if (count > 1)
                Console.Write("s");
            Console.WriteLine(" in the update queue.");
            Console.WriteLine("Updating database, please wait...");

            //Run add entry command
            AddEntry.Init();
            using (var db = new LiteDatabase(Database))
            {
                var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                foreach (var q in queue)
                {
                    var r = dbKeyRatio.FindOne(Query.Contains("Cik", q));
                    if (r == null)
                    {
                        Utility.WriteToLog(q + ": Could not find in database", ErrorLog);
                    }
                    else
                    {
                        AddEntry.Add(r.Ticker, false);
                    }
                }

                foreach (var amendment in amendments)
                {
                    var r = dbKeyRatio.FindOne(Query.Contains("Cik", amendment));
                    if (r != null)
                    {
                        AddEntry.Add(r.Ticker, true);
                    }
                }
            }
            Utility.PurgeTempDirectory();
            
            Console.WriteLine("Database updated!");
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
            /*
            Console.WriteLine("Verifying integrity...");
            var list = VerifyFiscalIntegrity();
            if (list.Count != 0)
            {
                var message = "The following ticker(s) lack financial info:\n";
                using (var db = new LiteDatabase(Database))
                {
                    var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                    message = list
                        .Select(l => dbKeyRatio.FindOne(Query.EQ("Cik", l)))
                        .Aggregate(message, (current, r) => current + r.Ticker + "\n");
                }
                Utility.WriteToLog(message, "FiscalIntegrityAnalysisError.txt");
                Console.WriteLine("Some companies are missing information.");
                Console.WriteLine("More information written to desktop");
            }
            else
            {
                Console.WriteLine("Database integrity analysis successful!");
            }
            */
        }

        private static bool DownloadFormsIndex()
        {
            Console.WriteLine("Checking for updates, please wait...");
            bool success;
            var attempt = 1;
            do
            {
                //Download file
                using (var wd = new Utility.WebDownload())
                {
                    try
                    {
                        wd.DownloadFile(SecFormAddress, Temp + "form.gz");
                    }
                    catch (Exception)
                    {
                        if (attempt > 3)
                        {
                            Console.WriteLine("Could not download index. Please try again later.");
                            Console.WriteLine("Press any key to continue...");
                            Console.ReadKey();
                            return false;
                        }
                        Console.WriteLine("Attempt " + attempt + " of 3 failed...");
                        attempt++;
                    }
                }
                success = true;
            } while (!success);

            //Decompress form.gz
            var file = new FileInfo(Temp + "form.gz");
            using (var originalFileStream = file.OpenRead())
            {
                var currentFileName = file.FullName;
                var newFileName = currentFileName.Remove(currentFileName.Length - file.Extension.Length);

                using (var decompressedFileStream = File.Create(newFileName + ".idx"))
                {
                    using (var decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                    {
                        decompressionStream.CopyTo(decompressedFileStream);
                    }
                }
            }
            return true;
        }

        private static List<string> ParseFormsIndex()
        {
            var filings = new List<string>();

            //Read lines of file individually, to cut RAM usage
            var index = new StreamReader(Temp + "form.idx");
            string line;
            char[] sep = { ' ', ' ' }; //To ensure that records are split, not just words

            while ((line = index.ReadLine()) != null)
            {
                //var tokens = line.Split(sep);
                var tokens = Regex.Split(line, @"\s{2,}"); //Split if string has more than 2 spaces
                tokens = tokens.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                try
                {
                    var form = tokens[0];
                    filings.AddRange(from s in Forms where form == s select tokens[2]);
                }
                catch (IndexOutOfRangeException)
                {
                    // ignored, it means the array is empty
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                }
            }
            return filings.Distinct().ToList();
        }

        private static List<string> Get10KAmendments()
        {
            var filings = new List<string>();

            //Read lines of file individually, to cut RAM usage
            var index = new StreamReader(Temp + "form.idx");
            string line;
            char[] sep = { ' ', ' ' }; //To ensure that records are split, not just words

            while ((line = index.ReadLine()) != null)
            {
                //var tokens = line.Split(sep);
                var tokens = Regex.Split(line, @"\s{2,}"); //Split if string has more than 2 spaces
                tokens = tokens.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                try
                {
                    var form = tokens[0];
                    filings.AddRange(from s in AmendForms where form == s select tokens[2]);
                }
                catch (IndexOutOfRangeException)
                {
                    // ignored, it means the array is empty
                }
                catch (Exception e)
                {
                    Utility.WriteToErrorLog(e);
                }
            }
            return filings.Distinct().ToList();
        }

        private static List<string> EndingFiscalYearCompanies()
        {
            var tickers = new List<string>();

            using (var db = new LiteDatabase(Database))
            {
                var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                var all = dbKeyRatio.Find(Query.All());

                //Get current month, along with previous 2 months
                //SEC requires all 10-K's be filed 60-90 days
                var date = Convert.ToDateTime(DateTime.Now.ToString(CultureInfo.CurrentCulture));
                var currentYear = int.Parse(date.Year.ToString());
                var currentMonth = GetMonthName(date.Month.ToString());
                var previousMonth = GetMonthName(date.AddMonths(-1).Month.ToString(CultureInfo.CurrentCulture));
                var twoMonths = GetMonthName(date.AddMonths(-2).Month.ToString(CultureInfo.CurrentCulture));
                var threeMonths = GetMonthName(date.AddMonths(-3).Month.ToString(CultureInfo.CurrentCulture));

                foreach (var keyRatio in all)
                {
                    if (keyRatio.FiscalYearEnd != currentMonth && keyRatio.FiscalYearEnd != previousMonth &&
                        keyRatio.FiscalYearEnd != twoMonths && keyRatio.FiscalYearEnd != threeMonths) continue;
                    //Checks if the company already contains current financial data
                    if (GetYear(keyRatio.FiscalYears[keyRatio.FiscalYears.Length - 1]) < currentYear)
                    {
                            tickers.Add(ParseCik(keyRatio.Cik));
                    }
                }
            }
            return tickers;
        }

        /// <summary>
        /// Verifies that all database entries contain all available financial information
        /// </summary>
        /// <returns>List of CIKs that lack financial information</returns>
        private static List<string> VerifyFiscalIntegrity()
        {
            var redownloadQueue = new List<string>();
            using (var db = new LiteDatabase(Database))
            {
                var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                var all = dbKeyRatio.Find(Query.All());

                redownloadQueue.AddRange(
                    from ratio 
                    in all
                    let list = ratio.FiscalYears.Select(GetYear).ToList()
                    where !list.Select((i, j) => i - j).Distinct().Skip(1).Any()
                    select ratio.Cik);
            }
            return redownloadQueue;
        }

        private static string GetMonthName(string date)
        {
            int.TryParse(date, out var month);
            return CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
        }

        //Gets a version of the CIK without leading zeros
        private static string ParseCik(string fullCik)
        {
            var subCik = "";
            if (fullCik == null)
            {
                return null;
            }
            var cik = fullCik.ToCharArray();
            var i = 0;
            while (true)
            {
                if (cik[i] != '0')
                    break;
                i++;
            }
            for (; i < cik.Length; i++)
            {
                subCik += cik[i];
            }
            return subCik;
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
