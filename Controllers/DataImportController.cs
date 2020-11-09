using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ImportDataFromExcel.Models;
using System.Runtime.InteropServices;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Text;
using System.Net.Http.Headers;
using System.Web.Configuration;
using System.Xml.Linq;
using System.Data;
using System.Web.Http.Routing;

namespace ImportDataFromExcel.Controllers
{
    public class DataImportController : Controller
    {
        //private string Username = "marilo@utdsoptimalchoice.com.uat";
        //private string Password = "Projekti123";
        //private string ClientId = "3MVG9c1ghSpUbLl.WS5lVK4WUp7.pJRpX9Stoq_maEArt4yFyRVoHQGlb_pTjebqcgaX6I0iWJY1ch7Mznkqw";
        //private string ClientSecret = "A4A69F3EFF63203E8D1CAE33B3C0A47717ED3950339AC9CE1F159C85B0F993E1";

        public const string ApiEndpoint = "/services/data/v36.0/";//"/services/data/00D030000008aiM/";
        public string LoginEndpoint = "";//"https://test.salesforce.com/services/oauth2/token"; //https://login.salesforce.com/services/oauth2/token
        public string AuthToken = "";
        public string ServiceUrl = "";
        private Excel.Application application = null;
        private Excel.Workbook workBook = null;
        private Excel.Worksheet workSheet = null;
        private string Status = "";
        private string Object = "";
        private int RecordCreated = 0;
        private int RecordFailed = 0;
        private DateTime StartDate = DateTime.Now;
        private double ProcessingTime = 0.0;
        private string MessageError = "";
        private string StatusCode = "";
        private string ReferenceId = "";

        static HttpClient Client;

        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelFile, FormCollection form)
        {
            try
            {
                LoginEndpoint = WebConfigurationManager.AppSettings["LoginEndpoint"];
                string Username = WebConfigurationManager.AppSettings["Username"];
                string Password = WebConfigurationManager.AppSettings["Password"];
                string ClientId = WebConfigurationManager.AppSettings["ClientId"];
                string ClientSecret = WebConfigurationManager.AppSettings["ClientSecret"];

                Client = new HttpClient();
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

                HttpContent content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    {"grant_type", "password"},
                    {"client_id", ClientId},
                    {"client_secret", ClientSecret},
                    {"username", Username},
                    {"password", Password}
                });

                HttpResponseMessage message = Client.PostAsync(LoginEndpoint, content).Result;

                string response = message.Content.ReadAsStringAsync().Result;
                JObject obj = JObject.Parse(response);

                AuthToken = (string)obj["access_token"];
                ServiceUrl = (string)obj["instance_url"];

                //curl https://yourInstance.salesforce.com/services/data/v34.0/composite/tree/Account/ -H "Authorization: Bearer token -H "Content-Type: application/json" -d "@newrecords.json"

                //string json = "{";
                //json += "\"records\" :[{";
                //json += "    \"attributes\" : {\"type\" : \"Account\", \"referenceId\" : \"Row1\"},";
                //json += "    \"name\" : \"SampleAccount1\",";
                //json += "    \"phone\" : \"1111111111\",";
                //json += "    \"website\" : \"www.salesforce.com\",";
                //json += "    \"industry\" : \"Banking\"";
                //json += "    },{";
                //json += "    \"attributes\" : { \"type\" : \"Account\", \"referenceId\" : \"Row2\"},";
                //json += "    \"name\" : \"SampleAccount2\",";
                //json += "    \"phone\" : \"2222222222\",";
                //json += "    \"website\" : \"www.salesforce2.com\",";
                //json += "    \"industry\" : \"Banking\"";
                //json += "    },{";
                //json += "    \"attributes\" : { \"type\" : \"Account\", \"referenceId\" : \"Row3\"},";
                //json += "    \"name\" : \"SampleAccount23\",";
                //json += "    \"phone\" : \"2222222222\",";
                //json += "    \"website\" : \"www.salesforce2.com\",";
                //json += "    \"industry\" : \"Banking\"";
                //json += "    },{";
                //json += "    \"attributes\" : { \"type\" : \"Account\", \"referenceId\" : \"Row4\"},";
                //json += "    \"name\" : \"SampleAccount4\",";
                //json += "    \"phone\" : \"2222222222\",";
                //json += "    \"website\" : \"www.salesforce2.com\",";
                //json += "    \"industry\" : \"Banking\"";
                //json += "    },{";
                //json += "    \"attributes\" : { \"type\" : \"Account\", \"referenceId\" : \"Row5\"},";
                //json += "    \"name\" : \"SampleAccount5\",";
                //json += "    \"phone\" : \"2222222222\",";
                //json += "    \"website\" : \"www.salesforce2.com\",";
                //json += "    \"industry\" : \"Banking\"";
                //json += "    }]";
                //json += "}";

                //Object = form["objectType"].ToString();
                ////string uri = $"" + ServiceUrl + "/services/data/v36.0/composite/tree/" + Object + "/";
                //string uri = $"" + ServiceUrl + "/services/data/v36.0/composite/tree/Account/";

                //HttpRequestMessage requestCreate = new HttpRequestMessage(HttpMethod.Post, uri);
                //requestCreate.Headers.Add("Authorization", "Bearer " + AuthToken);
                //requestCreate.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
                //requestCreate.Content = new StringContent(json, Encoding.UTF8, "application/json");

                //HttpResponseMessage response2 = Client.SendAsync(requestCreate).Result;
                //var temp = response2.Content.ReadAsStringAsync().Result;


                //CreateAccount(Client);
                //GetAccount();
                //CreateNewGasTariffs(Client);

                /*
                 //Electricity_Tariff_Price__c
                 1 Electricity Tariff ID = Electricity_Tariff__c
                 2 PES Area ID = PES_Area__c
                 3 Profile Class = Profile_Code__c
                 4 Tariff Type = Tariff_Type__c
                 5 Usage Band Min = Usage_Band_Min__c
                 6 Usage Band Max = Usage_Band_Max__c
                 7 Earliest Contract Start Date = EarliestContractStartDate__c
                 8 Latest Contract Start Date = LatestContractStartDate__c
                 9 Daily Standing Charge = Standing_Charge__c
                 10 Monthly Standing Charge = StandingChargeMonthly__c
                 11 Quarterly Standing Charge = StandingChargeQuarterly__c
                 12 Daily Standing Charge AMR = StandingChargeAMR__c
                 13 Monthly Standing Charge AMR = StandingChargeMonthlyAMR__c
                 14 Quarterly Standing Charge AMR = StandingChargeQuarterlyAMR__c
                 15 Unit Rate = Unit_Rate__c
                 16 Night Unit Rate = Night_Rate__c
                 17 Evening Weekend Unit Rate = Weekend_Rate__c
                 18 FIT Charge = FiTCharge__c
                 19 Pricing Start = Pricing_Start__c
                 20 Pricing End = Pricing_End__c
                */

                StartDate = DateTime.Now;

                if ((excelFile == null) || (excelFile.ContentLength == 0))
                {
                    ViewBag.Error = "Please select an excel file!";
                    return View("Index");
                }
                else
                {
                    if ((excelFile.FileName.EndsWith("xls")) || (excelFile.FileName.EndsWith("xlsx")) || (excelFile.FileName.EndsWith("csv")))
                    {
                        string path = Server.MapPath("~/Content/" + excelFile.FileName);
                        if (System.IO.File.Exists(path))
                            System.IO.File.Delete(path);
                        excelFile.SaveAs(path);

                        application = new Excel.Application();
                        workBook = application.Workbooks.Open(path);
                        workSheet = workBook.ActiveSheet;
                        Excel.Range range = workSheet.UsedRange;

                        bool isElectricityTariffPrice = true;
                        Object = form["objectType"].ToString();
                        if (!Object.Equals("Electricity_Tariff_Price__c"))
                            isElectricityTariffPrice = false;

                        string uri = $"" + ServiceUrl + "/services/data/v36.0/composite/tree/" + Object + "/";

                        HttpRequestMessage requestCreate = new HttpRequestMessage(HttpMethod.Post, uri);
                        requestCreate.Headers.Add("Authorization", "Bearer " + AuthToken);
                        requestCreate.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));

                        int multipleRecordCreateNo = 0;
                        HttpResponseMessage responseCreate = null;
                        XDocument doc = null;
                        string result = null;
                        int numberOfRows = range.Rows.Count;
                        string json = "{";
                        json += "\"records\" :[";

                        for (int row = 2; row <= 6; row++)
                        //for (int row = 2; row <= range.Rows.Count; row++)
                        {
                            multipleRecordCreateNo++;

                            json += "{";
                            json += "    \"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                            if(isElectricityTariffPrice)
                                json += CreateElectricityTariffPrice(range, row);
                            else
                                json += CreateGasTariffPrice(range, row);

                            if (json.Last() == ',')
                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                            json += "    },";

                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                            {
                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                json += "    ]";
                                json += "}";

                                requestCreate.Content = new StringContent(json, Encoding.UTF8, "application/json");
                                responseCreate = Client.SendAsync(requestCreate).Result;
                                result = responseCreate.Content.ReadAsStringAsync().Result;

                                doc = XDocument.Parse(result);
                                if (doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("hasErrors").ElementAt(0).Value.Equals("true"))
                                {
                                    ImportFailed(doc);
                                    return View("Error");
                                }

                                requestCreate = new HttpRequestMessage(HttpMethod.Post, uri);
                                requestCreate.Headers.Add("Authorization", "Bearer " + AuthToken);
                                requestCreate.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
                                json = "{";
                                json += "\"records\" :[";
                                RecordCreated += multipleRecordCreateNo;
                                multipleRecordCreateNo = 0;
                            }
                        }

                        json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                        json += "    ]";
                        json += "}";

                        requestCreate.Content = new StringContent(json, Encoding.UTF8, "application/json");

                        responseCreate = Client.SendAsync(requestCreate).Result;
                        result = responseCreate.Content.ReadAsStringAsync().Result;

                        doc = XDocument.Parse(result);
                        if (doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("hasErrors").ElementAt(0).Value.Equals("true"))
                        {
                            ImportFailed(doc);
                            return View("Error");
                        }

                        CloseExcelFile();

                        Status = "Completed";
                        RecordCreated = numberOfRows - 1;
                        PopulateOutputTable();

                        return View("Success");
                    }
                    else
                    {
                        ViewBag.Error = "File type is incorrect! <br>";
                        return View("Index");
                    }
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public string CreateElectricityTariffPrice(Excel.Range range, int row)
        {
            string json = "";

            if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                json += "    \"Electricity_Tariff__c\" : \"" + ((Excel.Range)range.Cells[row, 1]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                json += "    \"PES_Area__c\" : \"" + ((Excel.Range)range.Cells[row, 2]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                json += "    \"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 3]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                json += "    \"Tariff_Type__c\" : \"" + ((Excel.Range)range.Cells[row, 4]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                json += "    \"Usage_Band_Min__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                json += "    \"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 6]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                json += "    \"EarliestContractStartDate__c\" : \"" + ((Excel.Range)range.Cells[row, 7]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                json += "    \"LatestContractStartDate__c\" : \"" + ((Excel.Range)range.Cells[row, 8]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 9]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                json += "    \"StandingChargeMonthly__c\" : \"" + ((Excel.Range)range.Cells[row, 10]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                json += "    \"StandingChargeQuarterly__c\" : \"" + ((Excel.Range)range.Cells[row, 11]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                json += "    \"StandingChargeAMR__c\" : \"" + ((Excel.Range)range.Cells[row, 12]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                json += "    \"StandingChargeMonthlyAMR__c\" : \"" + ((Excel.Range)range.Cells[row, 13]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                json += "    \"StandingChargeQuarterlyAMR__c\" : \"" + ((Excel.Range)range.Cells[row, 14]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 15] != null) && (((Excel.Range)range.Cells[row, 15]).Text != string.Empty))
                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 15]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 16] != null) && (((Excel.Range)range.Cells[row, 16]).Text != string.Empty))
                json += "    \"Night_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 16]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 17] != null) && (((Excel.Range)range.Cells[row, 17]).Text != string.Empty))
                json += "    \"Weekend_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 17]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 18] != null) && (((Excel.Range)range.Cells[row, 18]).Text != string.Empty))
                json += "    \"FiTCharge__c\" : \"" + ((Excel.Range)range.Cells[row, 18]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 19] != null) && (((Excel.Range)range.Cells[row, 19]).Text != string.Empty))
                json += "    \"Pricing_Start__c\" : \"" + ((Excel.Range)range.Cells[row, 19]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 20] != null) && (((Excel.Range)range.Cells[row, 20]).Text != string.Empty))
                json += "    \"Pricing_End__c\" : \"" + ((Excel.Range)range.Cells[row, 20]).Text + "\"";

            return json;
        }

        public string CreateGasTariffPrice(Excel.Range range, int row)
        {
            string json = "";

            if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                json += "    \"Gas_Tariff__c\" : \"" + ((Excel.Range)range.Cells[row, 1]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                json += "    \"PES_Area__c\" : \"" + ((Excel.Range)range.Cells[row, 2]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                json += "    \"Usage_Band_Min__c\" : \"" + ((Excel.Range)range.Cells[row, 3]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                json += "    \"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 4]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                json += "    \"EarliestContractStartDate__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                json += "    \"LatestContractStartDate__c\" : \"" + ((Excel.Range)range.Cells[row, 6]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 7]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                json += "    \"StandingChargeMonthly__c\" : \"" + ((Excel.Range)range.Cells[row, 8]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                json += "    \"StandingChargeQuarterly__c\" : \"" + ((Excel.Range)range.Cells[row, 9]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                json += "    \"StandingChargeAMR__c\" : \"" + ((Excel.Range)range.Cells[row, 10]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                json += "    \"StandingChargeMonthlyAMR__c\" : \"" + ((Excel.Range)range.Cells[row, 11]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                json += "    \"StandingChargeQuarterlyAMR__c\" : \"" + ((Excel.Range)range.Cells[row, 12]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 13]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                json += "    \"Pricing_Start__c\" : \"" + ((Excel.Range)range.Cells[row, 14]).Text + "\",";
            if (((Excel.Range)range.Cells[row, 15] != null) && (((Excel.Range)range.Cells[row, 15]).Text != string.Empty))
                json += "    \"Pricing_End__c\" : \"" + ((Excel.Range)range.Cells[row, 15]).Text + "\",";

            return json;
        }

        public void ImportFailed(XDocument doc)
        {
            Status = "Failed";
            RecordFailed = 1;
            MessageError = doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("results").ElementAt(0).Descendants("errors").ElementAt(0).Descendants("message").ElementAt(0).Value;
            StatusCode = doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("results").ElementAt(0).Descendants("errors").ElementAt(0).Descendants("statusCode").ElementAt(0).Value;
            ReferenceId = doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("results").ElementAt(0).Descendants("referenceId").ElementAt(0).Value;

            CloseExcelFile();
            PopulateOutputTable();
        }

        public void PopulateOutputTable()
        {
            ProcessingTime = (DateTime.Now - StartDate).TotalSeconds;

            Results results = new Results();
            results.Status = Status;
            results.Object = Object;
            results.RecordCreated = RecordCreated.ToString();
            results.RecordFailed = RecordFailed.ToString();
            results.StartDate = StartDate.ToString();
            results.ProcessingTime = (Math.Round(ProcessingTime, 2)).ToString();
            results.MessageError = MessageError;
            results.StatusCode = StatusCode;
            results.ReferenceId = ReferenceId;
            ViewBag.Results = results;
        }

        public void CloseExcelFile()
        {
            workBook.Close(true, null, null);
            application.Quit();
            Marshal.ReleaseComObject(workSheet);
            Marshal.ReleaseComObject(workBook);
            Marshal.ReleaseComObject(application);
        }



        private void CreateAccount(HttpClient client)
        {
            string companyName = "Test123";
            string phone = "123-456-7890";

            string createMessage = 
                $"<root>" +
                    $"<Name>{companyName}</Name>" +
                    $"<Phone>{phone}</Phone>" +
                $"</root>";

            string result = CreateRecord(client, createMessage, "Account");

            XDocument doc = XDocument.Parse(result);

            string id = ((XElement)doc.Root.FirstNode).Value;
            string success = ((XElement)doc.Root.LastNode).Value;
        }

        private void CreateObjectType1(HttpClient client, Excel.Range range, int row)
        {

            string createMessage = $"<root>";

            if ((Excel.Range)range.Cells[row, 1] != null)
                createMessage += $"<Name>{ ((Excel.Range)range.Cells[row, 1]).Text }</Name>";
            
            createMessage += $"</root>";

            string result = CreateRecord(client, createMessage, "Account");

            XDocument doc = XDocument.Parse(result);

            string id = ((XElement)doc.Root.FirstNode).Value;
            string success = ((XElement)doc.Root.LastNode).Value;
        }

        public void GetAccount()
        {
            string name = "a0ba000000GdKld";
            string queryMessage = $"SELECT Id, Name, Tariff_Display_Name__c FROM Gas_Tariff__c WHERE Id = '{name}'";

            JObject obj = JObject.Parse(QueryRecord(Client, queryMessage));

            if ((string)obj["totalSize"] == "1")
            {
                // Only one record, use it
                string accountId = (string)obj["records"][0]["Id"];
                string accountPhone = (string)obj["records"][0]["Phone"];
            }
            if ((string)obj["totalSize"] == "0")
            {
                // No record, create an Account
            }
            else
            {
                // Multiple records, either filter further to determine correct Account or choose the first result
            }
        }

        private string QueryRecord(HttpClient client, string queryMessage)
        {
            string restQuery = $"{ServiceUrl}{ApiEndpoint}query?q={queryMessage}";

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, restQuery);
            request.Headers.Add("Authorization", "Bearer " + AuthToken);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            HttpResponseMessage response = client.SendAsync(request).Result;
            return response.Content.ReadAsStringAsync().Result;
        }

        public void CreateNewGasTariffs(HttpClient client)
        {
            string companyName = "Test123";
            string owner = GetUserId("UTDS Optimal Choice");

            string createMessage =
                $"<root>" +
                    $"<Tariff_Display_Name__c>{companyName}</Tariff_Display_Name__c>" +
                    $"<Tariff_Name__c>{companyName}</Tariff_Name__c>" +
                    $"<TariffActualDuration__c>{1}</TariffActualDuration__c>" +
                    $"<OwnerId>{owner}</OwnerId>" +
                $"</root>";

            string result = CreateRecord(client, createMessage, "Gas_Tariff__c");

            XDocument doc = XDocument.Parse(result);

            string id = ((XElement)doc.Root.FirstNode).Value;
            string success = ((XElement)doc.Root.LastNode).Value;
        }

        public string GetUserId(string name)
        {
            string queryMessage = $"SELECT Id, Name FROM User WHERE Name = '{name}'";

            JObject obj = JObject.Parse(QueryRecord(Client, queryMessage));

            if ((string)obj["totalSize"] == "1")
            {
                // Only one record, use it
                return (string)obj["records"][0]["Id"];
                //string accountPhone = (string)obj["records"][0]["Phone"];
            }
            if ((string)obj["totalSize"] == "0")
            {
                // No record, create an Account
            }
            else
            {
                // Multiple records, either filter further to determine correct Account or choose the first result
            }

            return "";
        }

        public void CreateNewElectricityTariff(HttpClient client, Excel.Range range, int row)
        {
            string createMessage = $"<root>";

            if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                createMessage += $"<Electricity_Tariff__c>{ ((Excel.Range)range.Cells[row, 1]).Text }</Electricity_Tariff__c>";
            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                createMessage += $"<PES_Area__c>{ ((Excel.Range)range.Cells[row, 2]).Text }</PES_Area__c>";
            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                createMessage += $"<Profile_Code__c>{ ((Excel.Range)range.Cells[row, 3]).Text }</Profile_Code__c>";
            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                createMessage += $"<Tariff_Type__c>{ ((Excel.Range)range.Cells[row, 4]).Text }</Tariff_Type__c>";
            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                createMessage += $"<Usage_Band_Min__c>{ ((Excel.Range)range.Cells[row, 5]).Text }</Usage_Band_Min__c>";
            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                createMessage += $"<Usage_Band_Max__c>{ ((Excel.Range)range.Cells[row, 6]).Text }</Usage_Band_Max__c>";
            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                createMessage += $"<EarliestContractStartDate__c>{ ((Excel.Range)range.Cells[row, 7]).Text }</EarliestContractStartDate__c>";
            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                createMessage += $"<LatestContractStartDate__c>{ ((Excel.Range)range.Cells[row, 8]).Text }</LatestContractStartDate__c>";
            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                createMessage += $"<Standing_Charge__c>{ ((Excel.Range)range.Cells[row, 9]).Text }</Standing_Charge__c>";
            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                createMessage += $"<StandingChargeMonthly__c>{ ((Excel.Range)range.Cells[row, 10]).Text }</StandingChargeMonthly__c>";
            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                createMessage += $"<StandingChargeQuarterly__c>{ ((Excel.Range)range.Cells[row, 11]).Text }</StandingChargeQuarterly__c>";
            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                createMessage += $"<StandingChargeAMR__c>{ ((Excel.Range)range.Cells[row, 12]).Text }</StandingChargeAMR__c>";
            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                createMessage += $"<StandingChargeMonthlyAMR__c>{ ((Excel.Range)range.Cells[row, 13]).Text }</StandingChargeMonthlyAMR__c>";
            if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                createMessage += $"<StandingChargeQuarterlyAMR__c>{ ((Excel.Range)range.Cells[row, 14]).Text }</StandingChargeQuarterlyAMR__c>";
            if (((Excel.Range)range.Cells[row, 15] != null) && (((Excel.Range)range.Cells[row, 15]).Text != string.Empty))
                createMessage += $"<Unit_Rate__c>{ ((Excel.Range)range.Cells[row, 15]).Text }</Unit_Rate__c>";
            if (((Excel.Range)range.Cells[row, 16] != null) && (((Excel.Range)range.Cells[row, 16]).Text != string.Empty))
                createMessage += $"<Night_Rate__c>{ ((Excel.Range)range.Cells[row, 16]).Text }</Night_Rate__c>";
            if (((Excel.Range)range.Cells[row, 17] != null) && (((Excel.Range)range.Cells[row, 17]).Text != string.Empty))
                createMessage += $"<Weekend_Rate__c>{ ((Excel.Range)range.Cells[row, 17]).Text }</Weekend_Rate__c>";
            if (((Excel.Range)range.Cells[row, 18] != null) && (((Excel.Range)range.Cells[row, 18]).Text != string.Empty))
                createMessage += $"<FiTCharge__c>{ ((Excel.Range)range.Cells[row, 18]).Text }</FiTCharge__c>";
            if (((Excel.Range)range.Cells[row, 19] != null) && (((Excel.Range)range.Cells[row, 19]).Text != string.Empty))
                createMessage += $"<Pricing_Start__c>{ ((Excel.Range)range.Cells[row, 19]).Text }</Pricing_Start__c>";
            if (((Excel.Range)range.Cells[row, 20] != null) && (((Excel.Range)range.Cells[row, 20]).Text != string.Empty))
                createMessage += $"<Pricing_End__c>{ ((Excel.Range)range.Cells[row, 20]).Text }</Pricing_End__c>";

            createMessage += $"</root>";

            string result = CreateRecord(client, createMessage, Object);

            XDocument doc = XDocument.Parse(result);

            string id = ((XElement)doc.Root.FirstNode).Value;
            string success = ((XElement)doc.Root.LastNode).Value;
            if (success.Equals("true"))
                RecordCreated++;
            else
                RecordFailed++;
        }

        private string CreateRecord(HttpClient client, string createMessage, string recordType)
        {
            HttpContent contentCreate = new StringContent(createMessage, Encoding.UTF8, "application/xml");
            string uri = $"{ServiceUrl}{ApiEndpoint}sobjects/{recordType}";

            HttpRequestMessage requestCreate = new HttpRequestMessage(HttpMethod.Post, uri);
            requestCreate.Headers.Add("Authorization", "Bearer " + AuthToken);
            requestCreate.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
            requestCreate.Content = contentCreate;

            HttpResponseMessage response = client.SendAsync(requestCreate).Result;
            return response.Content.ReadAsStringAsync().Result;
        }
    }
}