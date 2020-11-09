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

                        //for (int row = 2; row <= 6; row++)
                        for (int row = 2; row <= range.Rows.Count; row++)
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
    }
}