using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace ToExcel
{
    public static class Holiday
    {
        private static string token = "cmljYXJkby5zLmpvcmdlQGhvdG1haWwuY29tJmhhc2g9MTI1OTQ0MQ";
        private static string cityCode = "3170206"; // uberlandia
        private static string url = "https://api.calendario.com.br/?json=true&ano={0}&ibge={1}&token={2}";
        private static int currentYear = DateTime.Now.Year;
        private static List<HolidayDate> holidayList = new List<HolidayDate>();

        static Holiday()
        {
           GetHolidays();
        }

        public static bool IsHolidayDate(DateTime dateTime)
        {
            return holidayList.Where(r => DateTime.Parse(r.date).Date == dateTime.Date).ToList().Count() > 0;
        }

        public static string GetHolidayName(DateTime dateTime)
        {
            return holidayList.Where(r => DateTime.Parse(r.date).Date == dateTime.Date).First().name;
        }

        private static void GetHolidays()
        {
            string text;
            string textFile = currentYear + ".json";
            if (File.Exists(textFile))
            {
                text = File.ReadAllText(textFile);
            }
            else
            {
                WebRequest request = WebRequest.CreateHttp(String.Format(url, currentYear, cityCode, token));
                request.Method = "GET";
                request.Credentials = CredentialCache.DefaultCredentials;
                ((HttpWebRequest)request).UserAgent = "GetHolidayDays";

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode != HttpStatusCode.OK)
                {
                    text = "";
                    response.StatusCode.ToString();
                }
                else
                {
                    Stream receiveStream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(receiveStream);
                    text = reader.ReadToEnd().ToString();
                    System.IO.File.WriteAllText(textFile, text);
                    receiveStream.Close();
                }
            }
            if (!String.IsNullOrEmpty(text))
            {
                holidayList = JsonConvert.DeserializeObject<List<HolidayDate>>(text);
            }
        }
    }

    public class HolidayDate
    {
        public string date { get; set; }
        public string name { get; set; }
        public string link { get; set; }
        public string type { get; set; }
        public string description { get; set; }
        public object type_code { get; set; }
        public string raw_description { get; set; }
    }

}
