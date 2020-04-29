using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Operation_Yellow_Pointer.Client.Platform
{
    class HistoricalPrices
    {
        static void GetHistoricalPriceData(string ticker, string[] beginningDateParameters)
        {
            var localParameters = new object[7];
            var currentDate = DateTime.Now;
            localParameters[0] = ticker;
            localParameters[1] = beginningDateParameters[0];
            localParameters[2] = beginningDateParameters[1];
            localParameters[3] = beginningDateParameters[2];
            localParameters[4] = currentDate.Year.ToString();
            localParameters[5] = currentDate.Month.ToString();
            localParameters[6] = currentDate.Day.ToString();
            /*
             * {0} = Ticker
             * {1} = Beginning Year
             * {2} = "" Month
             * {3} = "" Day
             * {4} = Current Year
             * {5} = "" Month
             * {6} = "" Day
             */
            var globalAPI =
                "http://globalquote.morningstar.com/globalcomponent/RealtimeHistoricalStockData.ashx?ticker={0}&showVol=true&dtype=his&f=d&curry=USD&range={1}-{2}-{3}|{4}-{5}-{6}&isD=true&isS=true&hasF=true&ProdCode=DIRECT";
            var localCall = string.Format(globalAPI, localParameters);

            dynamic priceDataList;
            try
            {
                string json;
                using (var webDownload = new Utility.WebDownload())
                {
                    json = webDownload.DownloadString(localCall);
                }
                dynamic array = JsonConvert.DeserializeObject(json);
                priceDataList = array.PriceDataList[0];
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return;
            }

            //Contains historical prices
            var dataPoints = new List<double>();
            var dateIndices = new List<int>();

            var datapoints = priceDataList.Datapoints;
            var dateIndexs = priceDataList.DateIndexs;
            foreach (var dp in datapoints)
            {
                var closingPrice = double.Parse(JArray.FromObject(dp)[0].ToString());
                dataPoints.Add(closingPrice);
            }
            foreach (var di in dateIndexs)
            {
                var date = int.Parse(di.ToString());
                //Add 2 to the date
                //Morningstar gets the date and, for some reason, is the correct day minus 2
                //Adds two to correct the error
                dateIndices.Add(date + 2);
            }
        }
    }
}
