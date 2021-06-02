using System.Configuration;
using System.Collections.Specialized;
using System;
using ScrapySharp;
using HtmlAgilityPack;
using ScrapySharp.Network;
using System.Collections.Generic;
using ScrapySharp.Extensions;
using System.Net;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

namespace OSMProject
{
    class Program
    {
        //Variables
        private static ScrapingBrowser scrapingBrowser = new ScrapingBrowser();
        private static readonly string websiteURL = "https://srh.bankofchina.com/search/whpj/searchen.jsp";

        private static List<string> currencyValues;

        private static readonly string outputFilesDir = @"D:\CurrencyCSVOutput\";
        //Functions
        public static void getCurrencyValues()
        {
            Console.WriteLine("Getting currency values from " + websiteURL);
            currencyValues = new List<string>();
            WebPage webPage = scrapingBrowser.NavigateToPage(new Uri(websiteURL));
            HtmlNode html = webPage.Html;
            var values = html.CssSelect("option");
            foreach (var val in values)
            {
                currencyValues.Add(new string(val.InnerHtml.ToString()));
            }

            //Remove first list object with "Select the currency" 
            currencyValues.RemoveAt(0);

        }
        private static HtmlNode visitWebsite(string url)
        {
            Console.WriteLine("Open website " + url);
            WebPage webPage = scrapingBrowser.NavigateToPage(new Uri(url));
            HtmlNode html = webPage.Html;
            return html;
        }

        private static int getPageNumber(string htmlResp)
        {
            Regex regex = new Regex("sorry");
            //No records for chosen currency
            if (regex.IsMatch(htmlResp))
            {
                Console.WriteLine("No records!");
                return -1;
            }
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlResp);
            HtmlNodeCollection nodes = htmlDoc.DocumentNode.SelectNodes(".//form/input");
            if (nodes == null)
            {
                Console.WriteLine("Error, no table data found!");
                return -1;
            }
            int cntNode = 0;
            int cntNodeAttribute = 0;
            int pageNum = 0;

            foreach (HtmlNode node in nodes)
            {
                if (cntNode == 3)
                {
                    HtmlAttributeCollection attributes = node.Attributes;
                    cntNodeAttribute = 0;
                    foreach (var att in attributes)
                    {
                        if (cntNodeAttribute == 2)
                        {
                            pageNum = int.Parse(att.Value);
                            break;
                        }
                        cntNodeAttribute++;

                    }
                    break;
                }
                cntNode++;
            }


            return pageNum;
        }

        private static String formatOutputData(string htmlResp, int dataPosition)
        {
           return  htmlResp.PadRight(22, ' ');
           //For better formatting,different pad length could be chosen for different dataPosition

        }

        private static String parseHtmlResponse(string htmlResp, string currency, string date1, string date2, int pageNumber)
        {
            //WebException
            if (htmlResp.Equals(""))
                return "";

            Regex regex = new Regex("sorry");
            //No records for choosen currency
            if (regex.IsMatch(htmlResp))
            {
                Console.WriteLine("No records for currency " + currency + ", for chosen dates: " + date1 + ", " + date2);
                return "";
            }

            //Load html document from response string
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlResp);

            //Select table with results
            //HtmlNodeCollection has table rows
            HtmlNodeCollection nodes = htmlDoc.DocumentNode.SelectNodes(".//table[@bgcolor]/tr/td");
            if (nodes == null)
            {
                Console.WriteLine("Error, no table data found!");
                return "";
            }

            StringBuilder sbOutput = new StringBuilder();

            int cnt = 0;
            foreach (HtmlNode node in nodes)
            {
                //Do not write header data for pageNumber > 1
                if (pageNumber > 1 && cnt < 7) {
                    cnt++;
                    continue;
                }
                String currentHtmlString = "";
                if (node != null){
                    if (node.FirstChild != null)
                         currentHtmlString = node.FirstChild.InnerHtml;
                }
                   

                if (cnt % 7 == 0)  sbOutput.Append("\n");
           
                if (currentHtmlString.Equals(""))
                    sbOutput.Append("                    ");
                else sbOutput.Append(formatOutputData(currentHtmlString, cnt));

                cnt++;
            }

            return sbOutput.ToString();

        }

        private static String sendParameters(string val, string date1, string date2, int pageNum)
        {
            HttpWebRequest myReq = (HttpWebRequest)WebRequest.Create(websiteURL);
            Console.WriteLine("\n\nPOST request for currency: " + val);
            string post_data = "erectDate=" + date1 + "&nothing=" + date2 + "&pjname=" + val + "&page=" + pageNum;
            myReq.Method = "POST";
            myReq.ContentType = "application/x-www-form-urlencoded";
            byte[] bytedata = Encoding.UTF8.GetBytes(post_data);
            myReq.ContentLength = bytedata.Length;

            Stream requestStream = myReq.GetRequestStream();
            requestStream.Write(bytedata, 0, bytedata.Length);
            requestStream.Close();


            string response = "";
            try
            {
                HttpWebResponse myResp = (HttpWebResponse)myReq.GetResponse();
                Console.WriteLine("Response status: " + myResp.StatusDescription);

                using (Stream dataStream = myResp.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);
                    response = reader.ReadToEnd();
                }

                // Close the http response.
                myResp.Close();
            }
            catch (WebException e)
            {
                Console.WriteLine("Web Exception occured! " + "\n\nException Message :" + e.Message);
                return response;
            }

            return response;

        }
        public static void processCurrencyRequest(string val, string date1, string date2)
        {
            StringBuilder sb = new StringBuilder();

            string strFilePath = outputFilesDir + val + "_" + date1 + "_" + date2 + ".csv";
            String htmlResp;
            int currentPageNum = 1;
            int expectedPageNum = 1;
            while (true)
            {
                htmlResp = sendParameters(val, date1, date2, expectedPageNum);
                currentPageNum = getPageNumber(htmlResp);
                if (currentPageNum < 0)
                    return;
                if (currentPageNum != expectedPageNum)
                {
                    Console.WriteLine("No more pages for processing!");
                    break;
                }
                Console.WriteLine("Reading from page number " + currentPageNum);
                sb.Append(parseHtmlResponse(htmlResp, val, date1, date2, currentPageNum) + "\n");
                expectedPageNum = currentPageNum + 1;

            }
            Console.WriteLine("Writting to output file for currency " + val + ", for chosen dates: " + date1 + ", " + date2);
            File.WriteAllText(strFilePath, sb.ToString());

        }

        public static void Main(string[] args)
        {
            //string outputDirLocation;
            //val = System.Configuration.ConfigurationManager.AppSettings.Get("OutputDir);
           
            //Request for all currency values
            getCurrencyValues();
            

            //Send HTTPRequest for each of the parameters
            string date2 = DateTime.UtcNow.ToString("yyyy-MM-dd");
            string date1 = DateTime.Today.AddDays(-2).ToString("yyyy-MM-dd");
           
            currencyValues.ForEach(val => processCurrencyRequest(val, date1, date2));
            

        }
    }
}
