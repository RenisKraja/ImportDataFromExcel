using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
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
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Reflection;

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
        private int recordCreated = 0;

        static HttpClient Client;

        private SelectList suppliers = new SelectList(new[]
        {
            new { ID = "1", Name = "British Gas Lite" },
            new { ID = "2", Name = "British Gas REN" },
            new { ID = "3", Name = "British Gas ACQ" },
            new { ID = "4", Name = "Smartest Energy" },
            new { ID = "5", Name = "Valda Electricity" },
            new { ID = "6", Name = "EDF" },
            new { ID = "7", Name = "Gazprom REN" },
            new { ID = "8", Name = "Gazprom ACQ" },
            new { ID = "9", Name = "Npower" },
            new { ID = "10", Name = "Opus Energy REN" },
            new { ID = "11", Name = "Opus Energy ACQ" },
            new { ID = "12", Name = "Scottish Power" },
            new { ID = "13", Name = "SSE" },
            new { ID = "14", Name = "CNG" },
            new { ID = "15", Name = "Crown Gas & Power" },
            new { ID = "16", Name = "Dyce Energy REN" },
            new { ID = "17", Name = "Dyce Energy ACQ" },
            new { ID = "18", Name = "EON" },
        },
        "ID", "Name", 1);

        private SelectList objectType = new SelectList(new[]
        {
            new { ID = "Electricity_Tariff_Price__c", Name = "Electricity Tariff Price" },
            new { ID = "Gas_Tariff_Price__c", Name = "Gas Tariff Price" },
        },
        "ID", "Name", 1);

        public ActionResult Index()
        {
            ViewData["suppliers"] = suppliers;
            ViewData["objectType"] = objectType;

            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelFile, FormCollection form, SSE_Dates model)
        {
            try
            {
                Methods methods = new Methods();

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

                ViewData["suppliers"] = suppliers;
                ViewData["objectType"] = objectType;

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

                        int supplierNO = Convert.ToInt32(form["suppliers"].ToString());

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
                        string unitType = string.Empty;
                        string electricityTariffId = string.Empty;
                        string gasTariffId = string.Empty;
                        string earliestContractStartDate = string.Empty;
                        string latestContractStartDate = string.Empty;

                        switch (supplierNO)
                        {
                            case 1:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        int passToRowNO = 2;
                                        //for (int row = 2; row <= 8; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 7] == null) || (((Excel.Range)range.Cells[row, 7]).Text == string.Empty))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 12]).Text != "DD"))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                //23-11-2020
                                                string date = ((Excel.Range)range.Cells[row, 2]).Text;

                                                string day = date.Substring(0, 2);
                                                if (day.Last() == '-')
                                                {
                                                    day = day.Remove(day.Length - 1, 1);
                                                    day = "0" + day;
                                                }

                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1] + "/" + day + "/" + date.Substring(date.Length - 4), "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1] + "/" + day + "/" + date.Substring(date.Length - 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 3]).Text;

                                                string day = date.Substring(0, 2);
                                                if (day.Last() == '-')
                                                {
                                                    day = day.Remove(day.Length - 1, 1);
                                                    day = "0" + day;
                                                }

                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1] + "/" + day + "/" + date.Substring(date.Length - 4), "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1] + "/" + day + "/" + date.Substring(date.Length - 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                                json += "\"PES_Area__c\" : \"" + methods.GetPESAreaID(((Excel.Range)range.Cells[row, 4]).Text) + "\",";
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                electricityTariffId = methods.GetElectricityTariffIdBGL(((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 9]).Text);
                                                if (electricityTariffId != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariffId + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                            {
                                                if (int.TryParse(((Excel.Range)range.Cells[row, 10]).Text, out int output))
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"" + methods.GetUsageBandMin(Int32.Parse(((Excel.Range)range.Cells[row, 10]).Text)) + "\",";
                                                    json += "\"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 10]).Text + "\",";
                                                }
                                                else
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"0\",";
                                                    json += "\"Usage_Band_Max__c\" : \"0\",";
                                                }
                                            }

                                            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                                            {
                                                unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 13]).Text);
                                                if (unitType != string.Empty)
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row, 14]).Text + "\",";
                                            }

                                            for (int innerRow = row; innerRow <= range.Rows.Count; innerRow++)
                                            {
                                                if (methods.GetUniqueIdentifierBGL(range, innerRow) == methods.GetUniqueIdentifierBGL(range, innerRow + 1))
                                                {
                                                    if (((Excel.Range)range.Cells[innerRow + 1, 13] != null) && (((Excel.Range)range.Cells[innerRow + 1, 13]).Text != string.Empty) && ((Excel.Range)range.Cells[innerRow + 1, 14] != null) && (((Excel.Range)range.Cells[innerRow + 1, 14]).Text != string.Empty))
                                                    {
                                                        unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[innerRow + 1, 13]).Text);
                                                        if (unitType != string.Empty)
                                                        {
                                                            json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[innerRow + 1, 14]).Text + "\",";
                                                            passToRowNO++;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    break;
                                                }
                                            }


                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            json += "\"Tariff_Type__c\" : \"1\",";

                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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

                                            passToRowNO++;
                                        }
                                    }
                                    else
                                    {
                                        int passToRowNO = 2;
                                        //for (int row = 2; row <= 6; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 7] == null) || (((Excel.Range)range.Cells[row, 7]).Text == string.Empty))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 11]).Text != "DD"))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 2]).Text;
                                                if (methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) != string.Empty)
                                                {
                                                    string day = date.Substring(0, 2);
                                                    if (day.Last() == '-')
                                                    {
                                                        day = day.Remove(day.Length - 1, 1);
                                                        day = "0" + day;
                                                    }
                                                    string year = date.Substring(date.Length - 2);
                                                    if (year == "20")
                                                        year = "2020";
                                                    else if (year == "21")
                                                        year = "2021";

                                                    //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + year, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + year)).ToString("yyyy-MM-dd") + "\",";
                                                }
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 3]).Text;
                                                if (methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) != string.Empty)
                                                {
                                                    string day = date.Substring(0, 2);
                                                    if (day.Last() == '-')
                                                    {
                                                        day = day.Remove(day.Length - 1, 1);
                                                        day = "0" + day;
                                                    }
                                                    string year = date.Substring(date.Length - 2);
                                                    if (year == "20")
                                                        year = "2020";
                                                    else if (year == "21")
                                                        year = "2021";

                                                    //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + year, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + year)).ToString("yyyy-MM-dd") + "\",";
                                                }
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 4]).Text;
                                                json += "    \"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                            {
                                                electricityTariffId = methods.GetGasTariffIdBGL(((Excel.Range)range.Cells[row, 7]).Text + ((Excel.Range)range.Cells[row, 8]).Text);
                                                if (electricityTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + electricityTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                if (int.TryParse(((Excel.Range)range.Cells[row, 9]).Text, out int output))
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"" + methods.GetUsageBandMinGas(Int32.Parse(((Excel.Range)range.Cells[row, 9]).Text)) + "\",";
                                                    json += "\"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 9]).Text + "\",";
                                                }
                                                else
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"0\",";
                                                    json += "\"Usage_Band_Max__c\" : \"0\",";
                                                }
                                            }

                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                            {
                                                unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 12]).Text);
                                                if (unitType != string.Empty)
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row, 13]).Text + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row + 1, 12] != null) && (((Excel.Range)range.Cells[row + 1, 12]).Text != string.Empty) && ((Excel.Range)range.Cells[row + 1, 13] != null) && (((Excel.Range)range.Cells[row + 1, 13]).Text != string.Empty))
                                            {
                                                unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[row + 1, 12]).Text);
                                                if (unitType != string.Empty)
                                                {
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row + 1, 13]).Text + "\",";
                                                }
                                            }
                                            passToRowNO++;


                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            //json += "\"Tariff_Type__c\" : \"1\",";

                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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

                                            passToRowNO++;
                                        }
                                    }
                                    break;
                                }
                            case 2:
                            case 3:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        int passToRowNO = 2;
                                        //for (int row = 2; row <= 498; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 8] == null) || (((Excel.Range)range.Cells[row, 8]).Text == string.Empty))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 13]).Text != "DD"))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                                            {
                                                int counter = 1;
                                                bool isEmpty = false;

                                                unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 14]).Text);
                                                if (unitType == string.Empty)
                                                {
                                                    isEmpty = true;
                                                }

                                                for (int innerRow = row; innerRow <= range.Rows.Count; innerRow++)
                                                {
                                                    if (methods.GetUniqueIdentifierBG(range, innerRow) == methods.GetUniqueIdentifierBG(range, innerRow + 1))
                                                    {
                                                        counter++;
                                                        if (((Excel.Range)range.Cells[innerRow + 1, 14] != null) && (((Excel.Range)range.Cells[innerRow + 1, 14]).Text != string.Empty) && ((Excel.Range)range.Cells[innerRow + 1, 15] != null) && (((Excel.Range)range.Cells[innerRow + 1, 15]).Text != string.Empty))
                                                        {
                                                            unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[innerRow + 1, 14]).Text);
                                                            if (unitType == string.Empty)
                                                            {
                                                                isEmpty = true;
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        break;
                                                    }
                                                }

                                                if (isEmpty)
                                                {
                                                    passToRowNO += counter;
                                                    continue;
                                                }
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 2]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                                //earliestContractStartDate = ((Excel.Range)range.Cells[row, 2]).Text;
                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(earliestContractStartDate.Substring(3, 2) + "/" + earliestContractStartDate.Substring(0, 2) + "/" + earliestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 3]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                                //latestContractStartDate = ((Excel.Range)range.Cells[row, 3]).Text;
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(latestContractStartDate.Substring(3, 2) + "/" + latestContractStartDate.Substring(0, 2) + "/" + latestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                                json += "\"PES_Area__c\" : \"" + methods.GetPESAreaID(((Excel.Range)range.Cells[row, 4]).Text) + "\",";
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                electricityTariffId = methods.GetElectricityTariffIdBGL(((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 9]).Text);
                                                if (electricityTariffId != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariffId + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                            {
                                                if (int.TryParse(((Excel.Range)range.Cells[row, 11]).Text, out int output))
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"" + methods.GetUsageBandMin(Int32.Parse(((Excel.Range)range.Cells[row, 11]).Text)) + "\",";
                                                    json += "\"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 11]).Text + "\",";
                                                }
                                                else
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"0\",";
                                                    json += "\"Usage_Band_Max__c\" : \"0\",";
                                                }
                                            }

                                            if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 15] != null) && (((Excel.Range)range.Cells[row, 15]).Text != string.Empty))
                                            {
                                                unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 14]).Text);
                                                if (unitType != string.Empty)
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row, 15]).Text + "\",";
                                            }

                                            for (int innerRow = row; innerRow <= range.Rows.Count; innerRow++)
                                            {
                                                if (methods.GetUniqueIdentifierBG(range, innerRow) == methods.GetUniqueIdentifierBG(range, innerRow + 1))
                                                {
                                                    if (((Excel.Range)range.Cells[innerRow + 1, 14] != null) && (((Excel.Range)range.Cells[innerRow + 1, 14]).Text != string.Empty) && ((Excel.Range)range.Cells[innerRow + 1, 15] != null) && (((Excel.Range)range.Cells[innerRow + 1, 15]).Text != string.Empty))
                                                    {
                                                        unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[innerRow + 1, 14]).Text);
                                                        if (unitType != string.Empty)
                                                        {
                                                            json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[innerRow + 1, 15]).Text + "\",";
                                                            passToRowNO++;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    break;
                                                }
                                            }


                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            json += "\"Tariff_Type__c\" : \"1\",";


                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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

                                            passToRowNO++;
                                        }
                                    }
                                    else
                                    {
                                        int passToRowNO = 2;
                                        //for (int row = 2; row <= 6; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 7] == null) || (((Excel.Range)range.Cells[row, 7]).Text == string.Empty))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 11]).Text != "DD"))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 2]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");
                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 3]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 4]).Text;
                                                json += "    \"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                            {
                                                if (supplierNO == 2)
                                                    gasTariffId = methods.GetGasTariffIdBG(((Excel.Range)range.Cells[row, 7]).Text + ((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 10]).Text);
                                                else
                                                    gasTariffId = methods.GetGasTariffIdBG_DSC(((Excel.Range)range.Cells[row, 7]).Text + ((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 10]).Text);

                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                if (int.TryParse(((Excel.Range)range.Cells[row, 9]).Text, out int output))
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"" + methods.GetUsageBandMinGas(Int32.Parse(((Excel.Range)range.Cells[row, 9]).Text)) + "\",";
                                                    json += "\"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 9]).Text + "\",";
                                                }
                                                else
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"0\",";
                                                    json += "\"Usage_Band_Max__c\" : \"0\",";
                                                }
                                            }

                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                            {
                                                unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 12]).Text);
                                                if (unitType != string.Empty)
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row, 13]).Text + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row + 1, 12] != null) && (((Excel.Range)range.Cells[row + 1, 12]).Text != string.Empty) && ((Excel.Range)range.Cells[row + 1, 13] != null) && (((Excel.Range)range.Cells[row + 1, 13]).Text != string.Empty))
                                            {
                                                unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[row + 1, 12]).Text);
                                                if (unitType != string.Empty)
                                                {
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row + 1, 13]).Text + "\",";
                                                }
                                            }
                                            passToRowNO++;


                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            //json += "\"Tariff_Type__c\" : \"1\",";

                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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

                                            passToRowNO++;
                                        }
                                    }
                                    break;
                                }
                            case 4:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        //for (int row = 3; row <= 6; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (((Excel.Range)range.Cells[row, 6] == null) || (((Excel.Range)range.Cells[row, 6]).Text == string.Empty))
                                                continue;

                                            if (
                                                ((Excel.Range)range.Cells[row, 4] != null)
                                                &&
                                                (((Excel.Range)range.Cells[row, 4]).Text != string.Empty)
                                                &&
                                                ((((Excel.Range)range.Cells[row, 4]).Text == "OP") || (((Excel.Range)range.Cells[row, 4]).Text.Substring(0, 2) == "HH"))
                                               )
                                            {
                                                continue;
                                            }
                                            if (
                                                ((Excel.Range)range.Cells[row, 6] != null)
                                                &&
                                                (((Excel.Range)range.Cells[row, 6]).Text != string.Empty)
                                                &&
                                                (((Excel.Range)range.Cells[row, 6]).Text.ToLower().Contains("level 2"))
                                               )
                                            {
                                                continue;
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                                json += "\"PES_Area__c\" : \"" + methods.GetPESAreaID(((Excel.Range)range.Cells[row, 2]).Text) + "\",";
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";

                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            {
                                                string electricityTariff = methods.GetElectricityTariffIdSE(((Excel.Range)range.Cells[row, 6]).Text);
                                                if (electricityTariff != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariff + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 7]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 8]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                                json += "    \"Night_Rate__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 9]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                                json += "    \"Weekend_Rate__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 10]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + ((Excel.Range)range.Cells[row, 11]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 12]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 13]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                json += "    \"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                                //json += "    \"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 14]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                json += "    \"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                                //json += "    \"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                            }

                                            json += "    \"Electricity_Tariff__c\" : \"a0h1B00000FLP7H\",";
                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            json += "\"Tariff_Type__c\" : \"1\",";


                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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
                                    }
                                    else
                                    {
                                        //for (int row = 4; row <= 5; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (((Excel.Range)range.Cells[row, 2] == null) || (((Excel.Range)range.Cells[row, 2]).Text == string.Empty))
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty) && (!((Excel.Range)range.Cells[row, 2]).Text.Contains("Level1")))
                                                continue;


                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                gasTariffId = methods.GetGasTariffIdSE(((Excel.Range)range.Cells[row, 3]).Text + ((Excel.Range)range.Cells[row, 2]).Text);
                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 4]).Text;
                                                json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 5]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 6]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 7]).Text) + "\",";
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 8]).Text) + "\",";

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            //json += "\"Tariff_Type__c\" : \"1\",";

                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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
                                    }
                                    break;
                                }
                            case 5:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        int passToRowNO = 2;
                                        //for (int row = 2; row <= 29; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 5] == null) || (((Excel.Range)range.Cells[row, 5]).Text == string.Empty))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 2]).Text.Contains("HH")))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 4]).Text.ToLower() == "off-peak"))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            for (int yearRow = 1; yearRow <= 3; yearRow++)
                                            {
                                                recordCreated++;
                                                multipleRecordCreateNo++;

                                                json += "{";
                                                json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "_" + yearRow + "\"},";

                                                if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                                                {
                                                    string pesArea = methods.GetPESAreaID(((Excel.Range)range.Cells[row, 1]).Text);
                                                    if (pesArea != string.Empty)
                                                        json += "\"PES_Area__c\" : \"" + pesArea + "\",";
                                                }
                                                if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                                    json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 2]).Text + "\",";

                                                if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                {
                                                    unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 5]).Text);
                                                    if (unitType != string.Empty)
                                                    {
                                                        if (yearRow == 1)
                                                        {
                                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                                            {
                                                                json += "\"" + unitType + "\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 6]).Text) + "\",";
                                                                json += "\"Electricity_Tariff__c\" : \"a0h1B00000ZkYvZ\",";
                                                            }
                                                        }
                                                        else if (yearRow == 2)
                                                        {
                                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                                            {
                                                                json += "\"" + unitType + "\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 7]).Text) + "\",";
                                                                json += "\"Electricity_Tariff__c\" : \"a0h1B00000ZkYve\",";
                                                            }
                                                        }
                                                        else if (yearRow == 3)
                                                        {
                                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                                            {
                                                                json += "\"" + unitType + "\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 8]).Text) + "\",";
                                                                json += "\"Electricity_Tariff__c\" : \"a0h1B00000ZkYvj\",";
                                                            }
                                                        }
                                                    }
                                                }

                                                for (int innerRow = row; innerRow <= range.Rows.Count; innerRow++)
                                                {
                                                    if (methods.GetUniqueIdentifierVE(range, innerRow) == methods.GetUniqueIdentifierVE(range, innerRow + 1))
                                                    {
                                                        unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[innerRow + 1, 5]).Text);
                                                        if (unitType != string.Empty)
                                                        {
                                                            if (yearRow == 1)
                                                            {
                                                                if (((Excel.Range)range.Cells[innerRow + 1, 6] != null) && (((Excel.Range)range.Cells[innerRow + 1, 6]).Text != string.Empty))
                                                                {
                                                                    json += "\"" + unitType + "\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[innerRow + 1, 6]).Text) + "\",";
                                                                }
                                                            }
                                                            else if (yearRow == 2)
                                                            {
                                                                if (((Excel.Range)range.Cells[innerRow + 1, 7] != null) && (((Excel.Range)range.Cells[innerRow + 1, 7]).Text != string.Empty))
                                                                {
                                                                    json += "\"" + unitType + "\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[innerRow + 1, 7]).Text) + "\",";
                                                                }
                                                            }
                                                            else if (yearRow == 3)
                                                            {
                                                                if (((Excel.Range)range.Cells[innerRow + 1, 8] != null) && (((Excel.Range)range.Cells[innerRow + 1, 8]).Text != string.Empty))
                                                                {
                                                                    json += "\"" + unitType + "\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[innerRow + 1, 8]).Text) + "\",";
                                                                }
                                                                passToRowNO++;
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        break;
                                                    }
                                                }

                                                json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                                json += "\"Tariff_Type__c\" : \"1\",";

                                                if (json.Last() == ',')
                                                    json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                                json += "},";

                                                if (yearRow == 3)
                                                {
                                                    if ((multipleRecordCreateNo == 198) && (row != range.Rows.Count))
                                                    {
                                                        json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                        json += "]";
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

                                                    passToRowNO++;
                                                }
                                            }


                                            //json += "{";
                                            //json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            //if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                                            //{
                                            //    string pesArea = GetPESAreaID(((Excel.Range)range.Cells[row, 1]).Text);
                                            //    if (pesArea != string.Empty)
                                            //        json += "\"PES_Area__c\" : \"" + pesArea + "\",";
                                            //}
                                            //if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            //    json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 2]).Text + "\",";

                                            //if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            //{
                                            //    unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 5]).Text);
                                            //    if (unitType != string.Empty)
                                            //    {
                                            //        double finalRate = 0.0;
                                            //        if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            //            finalRate += Convert.ToDouble(((Excel.Range)range.Cells[row, 6]).Text);
                                            //        if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                            //            finalRate += Convert.ToDouble(((Excel.Range)range.Cells[row, 7]).Text);
                                            //        if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                            //            finalRate += Convert.ToDouble(((Excel.Range)range.Cells[row, 8]).Text);

                                            //        json += "\"" + unitType + "\" : \"" + finalRate + "\",";
                                            //    }
                                            //}

                                            //for (int innerRow = row; innerRow <= range.Rows.Count; innerRow++)
                                            //{
                                            //    if (GetUniqueIdentifierVE(range, innerRow) == GetUniqueIdentifierVE(range, innerRow + 1))
                                            //    {
                                            //        if (((Excel.Range)range.Cells[innerRow + 1, 5] != null) && (((Excel.Range)range.Cells[innerRow + 1, 5]).Text != string.Empty) && ((Excel.Range)range.Cells[innerRow + 1, 6] != null) && (((Excel.Range)range.Cells[innerRow + 1, 6]).Text != string.Empty))
                                            //        {
                                            //            unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[innerRow + 1, 5]).Text);
                                            //            if (unitType != string.Empty)
                                            //            {
                                            //                json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[innerRow + 1, 6]).Text + "\",";

                                            //                double finalRate = 0.0;
                                            //                if (((Excel.Range)range.Cells[row + 1, 6] != null) && (((Excel.Range)range.Cells[row + 1, 6]).Text != string.Empty))
                                            //                    finalRate += Convert.ToDouble(((Excel.Range)range.Cells[row + 1, 6]).Text);
                                            //                if (((Excel.Range)range.Cells[row + 1, 7] != null) && (((Excel.Range)range.Cells[row + 1, 7]).Text != string.Empty))
                                            //                    finalRate += Convert.ToDouble(((Excel.Range)range.Cells[row + 1, 7]).Text);
                                            //                if (((Excel.Range)range.Cells[row + 1, 8] != null) && (((Excel.Range)range.Cells[row + 1, 8]).Text != string.Empty))
                                            //                    finalRate += Convert.ToDouble(((Excel.Range)range.Cells[row + 1, 8]).Text);

                                            //                json += "\"" + unitType + "\" : \"" + finalRate + "\",";


                                            //                passToRowNO++;
                                            //            }
                                            //        }
                                            //    }
                                            //    else
                                            //    {
                                            //        break;
                                            //    }
                                            //}


                                            //json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            //json += "\"Tariff_Type__c\" : \"1\",";

                                            //if (json.Last() == ',')
                                            //    json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            //json += "},";

                                            //if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            //{
                                            //    json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                            //    json += "]";
                                            //    json += "}";

                                            //    requestCreate.Content = new StringContent(json, Encoding.UTF8, "application/json");
                                            //    responseCreate = Client.SendAsync(requestCreate).Result;
                                            //    result = responseCreate.Content.ReadAsStringAsync().Result;

                                            //    doc = XDocument.Parse(result);
                                            //    if (doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("hasErrors").ElementAt(0).Value.Equals("true"))
                                            //    {
                                            //        ImportFailed(doc);
                                            //        return View("Error");
                                            //    }

                                            //    requestCreate = new HttpRequestMessage(HttpMethod.Post, uri);
                                            //    requestCreate.Headers.Add("Authorization", "Bearer " + AuthToken);
                                            //    requestCreate.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
                                            //    json = "{";
                                            //    json += "\"records\" :[";
                                            //    RecordCreated += multipleRecordCreateNo;
                                            //    multipleRecordCreateNo = 0;
                                            //}

                                            //passToRowNO++;
                                        }
                                    }
                                    else
                                    {
                                        ObjectDoesNotExist();
                                        return View("Error");
                                    }
                                    break;
                                }
                            case 6:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        //for (int row = 3; row <= 10; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (((Excel.Range)range.Cells[row, 5] == null) || (((Excel.Range)range.Cells[row, 5]).Text == string.Empty))
                                                continue;

                                            if (
                                                ((Excel.Range)range.Cells[row, 3] != null)
                                                &&
                                                (((Excel.Range)range.Cells[row, 3]).Text != string.Empty)
                                                &&
                                                ((((Excel.Range)range.Cells[row, 3]).Text.Contains("OUN")))
                                               )
                                            {
                                                continue;
                                            }

                                            if (
                                                ((Excel.Range)range.Cells[row, 11] != null)
                                                &&
                                                (((Excel.Range)range.Cells[row, 11]).Text != string.Empty)
                                                &&
                                                (!((((Excel.Range)range.Cells[row, 11]).Text == "0") || (((Excel.Range)range.Cells[row, 11]).Text == "0.0")))
                                               )
                                            {
                                                continue;
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                                            {
                                                string[] UB = ((Excel.Range)range.Cells[row, 1]).Text.Split('-');
                                                if (UB.Length == 2)
                                                {
                                                    json += "    \"Usage_Band_Min__c\" : \"" + UB[0] + "\",";
                                                    json += "    \"Usage_Band_Max__c\" : \"" + UB[1] + "\",";
                                                }
                                            }
                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                string region = ((Excel.Range)range.Cells[row, 2]).Text;
                                                json += "\"PES_Area__c\" : \"" + methods.GetPESAreaID(region.Substring(region.Length - 2)) + "\",";
                                            }


                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string profileClass = methods.GetProfileClassEDF(((Excel.Range)range.Cells[row, 3]).Text);
                                                if (profileClass != string.Empty)
                                                    json += "\"Profile_Code__c\" : \"" + profileClass + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                string electricityTariff = methods.GetElectricityTariffIdEDF(((Excel.Range)range.Cells[row, 5]).Text);
                                                if (electricityTariff != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariff + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 7]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 8]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                                json += "    \"Night_Rate__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 9]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                                json += "    \"Weekend_Rate__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 10]).Text) * 100) + "\",";

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            json += "\"Tariff_Type__c\" : \"1\",";


                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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
                                    }
                                    else
                                    {
                                        int passToRowNO = 2;
                                        //for (int row = 2; row <= 4; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 5] == null) || (((Excel.Range)range.Cells[row, 5]).Text == string.Empty))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 9]).Text != "0.0"))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                                            {
                                                string[] usageBand = ((Excel.Range)range.Cells[row, 1]).Text.Split('-');
                                                if (usageBand.Length == 2)
                                                    if (int.TryParse(usageBand[0], out int outputMin) && int.TryParse(usageBand[1], out int outputMax))
                                                    {
                                                        json += "\"Usage_Band_Min__c\" : \"" + usageBand[0] + "\",";
                                                        json += "\"Usage_Band_Max__c\" : \"" + usageBand[1] + "\",";
                                                    }
                                                    else
                                                    {
                                                        json += "\"Usage_Band_Min__c\" : \"0\",";
                                                        json += "\"Usage_Band_Max__c\" : \"0\",";
                                                    }
                                            }
                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 2]).Text;
                                                ldz = ldz.Substring(ldz.Length - 2);
                                                json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                gasTariffId = methods.GetGasTariffIdEDF(((Excel.Range)range.Cells[row, 5]).Text);
                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 7]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 8]).Text + "\",";

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            //json += "\"Tariff_Type__c\" : \"1\",";

                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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

                                            passToRowNO++;
                                        }
                                    }
                                    break;
                                }
                            case 7:
                            case 8:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        //for (int row = 3; row <= 6; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (((Excel.Range)range.Cells[row, 4] == null) || (((Excel.Range)range.Cells[row, 4]).Text == string.Empty))
                                                continue;

                                            // skip record if Unit Rate is empty
                                            if ((((Excel.Range)range.Cells[row, 8] == null) || (((Excel.Range)range.Cells[row, 8]).Text == string.Empty)) && (((Excel.Range)range.Cells[row, 9] == null) || (((Excel.Range)range.Cells[row, 9]).Text == string.Empty)))
                                                continue;

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 2]).Text + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                string electricityTariff = methods.GetElectricityTariffIdGazprom(((Excel.Range)range.Cells[row, 4]).Text);
                                                if (electricityTariff != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariff + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            {
                                                string pesArea = methods.GetPESAreaID(((Excel.Range)range.Cells[row, 6]).Text);
                                                if (pesArea != string.Empty)
                                                    json += "\"PES_Area__c\" : \"" + pesArea + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 7]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 8]).Text + "\",";
                                            else if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 9]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                                json += "    \"Night_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 10]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                                json += "    \"Weekend_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 11]).Text + "\",";

                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 5]).Text;
                                                if (date.ToLower().Contains("before"))
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                                    date = date.Substring(date.Length - 11); //31-Jan-2021
                                                    if (methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) != string.Empty)
                                                    {
                                                        string day = date.Substring(0, 2);
                                                        if (day.Last() == '-')
                                                        {
                                                            day = day.Remove(day.Length - 1, 1);
                                                            day = "0" + day;
                                                        }

                                                        json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + date.Substring(date.Length - 4))).ToString("yyyy-MM-dd") + "\",";
                                                        //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + date.Substring(date.Length - 4), "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                    }
                                                }
                                                else if (date.ToLower().Contains("after"))
                                                {
                                                    date = date.Substring(date.Length - 11); //31-Jan-2021
                                                    if (methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) != string.Empty)
                                                    {
                                                        string day = date.Substring(0, 2);
                                                        if (day.Last() == '-')
                                                        {
                                                            day = day.Remove(day.Length - 1, 1);
                                                            day = "0" + day;
                                                        }

                                                        json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + date.Substring(date.Length - 4))).ToString("yyyy-MM-dd") + "\",";
                                                        //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + date.Substring(date.Length - 4), "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                        json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + date.Substring(date.Length - 4)).AddMonths(6)).ToString("yyyy-MM-dd") + "\",";
                                                        //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(methods.GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + date.Substring(date.Length - 4), "MM/dd/yyyy", null).AddMonths(6)).ToString("yyyy-MM-dd") + "\",";
                                                    }
                                                }
                                            }

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            json += "\"Tariff_Type__c\" : \"1\",";


                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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
                                    }
                                    else
                                    {
                                        //for (int row = 4; row <= 5; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (((Excel.Range)range.Cells[row, 2] == null) || (((Excel.Range)range.Cells[row, 2]).Text == string.Empty))
                                                continue;

                                            int updateColumn = 0;
                                            // if first column is not empty
                                            if (((Excel.Range)range.Cells[row, 1]).Text != string.Empty)
                                                updateColumn = -1;

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2 + updateColumn] != null) && (((Excel.Range)range.Cells[row, 2 + updateColumn]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 6 + updateColumn] != null) && (((Excel.Range)range.Cells[row, 6 + updateColumn]).Text != string.Empty))
                                            {
                                                if (supplierNO == 7)
                                                    gasTariffId = methods.GetGasTariffIdGP_REN(((Excel.Range)range.Cells[row, 2 + updateColumn]).Text + ((Excel.Range)range.Cells[row, 6 + updateColumn]).Text);
                                                else
                                                    gasTariffId = methods.GetGasTariffIdGP_ACQ(((Excel.Range)range.Cells[row, 2 + updateColumn]).Text + ((Excel.Range)range.Cells[row, 6 + updateColumn]).Text);


                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4 + updateColumn] != null) && (((Excel.Range)range.Cells[row, 4 + updateColumn]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + ((Excel.Range)range.Cells[row, 4 + updateColumn]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 5 + updateColumn] != null) && (((Excel.Range)range.Cells[row, 5 + updateColumn]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 5 + updateColumn]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 7 + updateColumn] != null) && (((Excel.Range)range.Cells[row, 7 + updateColumn]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 7 + updateColumn]).Text;
                                                json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 8 + updateColumn] != null) && (((Excel.Range)range.Cells[row, 8 + updateColumn]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 8 + updateColumn]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                //if (date.Contains("-"))
                                                //    json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(methods.GetMonth(date.Substring(date.Length - 2) + "/" + date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + date.Substring(0, 4), "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                //else
                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 9 + updateColumn] != null) && (((Excel.Range)range.Cells[row, 9 + updateColumn]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 9 + updateColumn]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                //if (date.Contains("-"))
                                                //    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(methods.GetMonth(date.Substring(date.Length - 2) + "/" + date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + date.Substring(0, 4), "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                //else
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 10 + updateColumn] != null) && (((Excel.Range)range.Cells[row, 10 + updateColumn]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 10 + updateColumn]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 11 + updateColumn] != null) && (((Excel.Range)range.Cells[row, 11 + updateColumn]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 11 + updateColumn]).Text + "\",";

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            //json += "\"Tariff_Type__c\" : \"1\",";

                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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
                                    }
                                    break;
                                }
                            case 9:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        //for (int row = 3; row <= 32; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (((Excel.Range)range.Cells[row, 6] == null) || (((Excel.Range)range.Cells[row, 6]).Text == string.Empty))
                                                continue;

                                            // skip record if Unit Rate is empty
                                            if ((((Excel.Range)range.Cells[row, 12] == null) || (((Excel.Range)range.Cells[row, 12]).Text == string.Empty)) && (((Excel.Range)range.Cells[row, 13] == null) || (((Excel.Range)range.Cells[row, 13]).Text == string.Empty)))
                                                continue;

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                string pesArea = methods.GetPESAreaID(((Excel.Range)range.Cells[row, 2]).Text);
                                                if (pesArea != string.Empty)
                                                    json += "\"PES_Area__c\" : \"" + pesArea + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 3]).Text + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                string month = methods.GetMonth(((Excel.Range)range.Cells[row, 4]).Text);
                                                if (month != string.Empty)
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(month + "/01/" + ((Excel.Range)range.Cells[row, 5]).Text)).ToString("yyyy-MM-dd") + "\",";
                                                    //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(month + "/01/" + ((Excel.Range)range.Cells[row, 5]).Text, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(month + "/" + DateTime.DaysInMonth(Convert.ToInt32(((Excel.Range)range.Cells[row, 5]).Text), Convert.ToInt32(month)) + "/" + ((Excel.Range)range.Cells[row, 5]).Text).ToString("yyyy-MM-dd")) + "\",";
                                                    //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(month + "/" + DateTime.DaysInMonth(Convert.ToInt32(((Excel.Range)range.Cells[row, 5]).Text), Convert.ToInt32(month)) + "/" + ((Excel.Range)range.Cells[row, 5]).Text, "MM/dd/yyyy", null).ToString("yyyy-MM-dd")) + "\",";
                                                }
                                            }
                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            {
                                                string electricityTariff = methods.GetElectricityTariffIdNpower(((Excel.Range)range.Cells[row, 6]).Text);
                                                if (electricityTariff != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariff + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + ((Excel.Range)range.Cells[row, 9]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 10]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 11]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 12]).Text + "\",";
                                            else if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 13]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                                                json += "    \"Night_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 14]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 15] != null) && (((Excel.Range)range.Cells[row, 15]).Text != string.Empty))
                                                json += "    \"Weekend_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 15]).Text + "\",";
                                            else if (((Excel.Range)range.Cells[row, 16] != null) && (((Excel.Range)range.Cells[row, 16]).Text != string.Empty))
                                                json += "    \"Weekend_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 16]).Text + "\",";

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            json += "\"Tariff_Type__c\" : \"1\",";


                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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
                                    }
                                    else
                                    {
                                        //for (int row = 3; row <= 4; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (((Excel.Range)range.Cells[row, 10] == null) || (((Excel.Range)range.Cells[row, 10]).Text == string.Empty))
                                                continue;

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 2]).Text;
                                                json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                string month = methods.GetMonth(((Excel.Range)range.Cells[row, 4]).Text);
                                                if (month != string.Empty)
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(month + "/01/" + ((Excel.Range)range.Cells[row, 5]).Text)).ToString("yyyy-MM-dd") + "\",";
                                                    //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(month + "/01/" + ((Excel.Range)range.Cells[row, 5]).Text, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(month + "/" + DateTime.DaysInMonth(Convert.ToInt32(((Excel.Range)range.Cells[row, 5]).Text), Convert.ToInt32(month)) + "/" + ((Excel.Range)range.Cells[row, 5]).Text, "MM/dd/yyyy", null).ToString("yyyy-MM-dd")) + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(month + "/" + DateTime.DaysInMonth(Convert.ToInt32(((Excel.Range)range.Cells[row, 5]).Text), Convert.ToInt32(month)) + "/" + ((Excel.Range)range.Cells[row, 5]).Text).ToString("yyyy-MM-dd")) + "\",";
                                                }
                                            }
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + ((Excel.Range)range.Cells[row, 8]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + Decimal.Floor(Convert.ToDecimal(((Excel.Range)range.Cells[row, 9]).Text)) + "\",";
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                            {
                                                gasTariffId = methods.GetGasTariffIdNpower(((Excel.Range)range.Cells[row, 10]).Text);
                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 11]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 12]).Text + "\",";

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            //json += "\"Tariff_Type__c\" : \"1\",";

                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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
                                    }
                                    break;
                                }
                            case 10:
                            case 11:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        int passToRowNO = 2;
                                        //for (int row = 2; row <= 7; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 1] == null) || (((Excel.Range)range.Cells[row, 1]).Text == string.Empty))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty) &&
                                                ((((Excel.Range)range.Cells[row, 4]).Text.ToLower() == "off peak") || (((Excel.Range)range.Cells[row, 4]).Text.ToLower() == "hh") || (((Excel.Range)range.Cells[row, 4]).Text.ToLower() == "hh no availability") || (((Excel.Range)range.Cells[row, 4]).Text.ToLower() == "night & day") || (((Excel.Range)range.Cells[row, 4]).Text.ToLower() == "night saver"))
                                               )
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                                            {
                                                string profileClass = ((Excel.Range)range.Cells[row, 1]).Text;
                                                json += "\"PES_Area__c\" : \"" + methods.GetPESAreaID(profileClass.Substring(0, 2)) + "\",";
                                                json += "\"Profile_Code__c\" : \"" + profileClass.Substring(2, 1) + "\",";

                                                if (supplierNO == 10)
                                                    electricityTariffId = methods.GetElectricityTariffIdOE_REN(profileClass);
                                                else
                                                    electricityTariffId = methods.GetElectricityTariffIdOE_ACQ(profileClass);

                                                if (electricityTariffId != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                string unitValue = ((Excel.Range)range.Cells[row, 5]).Text;

                                                if (unitValue.IndexOf("kwh", StringComparison.Ordinal) > 0)
                                                {
                                                    unitValue = unitValue.Substring(0, unitValue.IndexOf("kwh", StringComparison.Ordinal));
                                                    if (!unitValue.Equals(string.Empty))
                                                        json += "\"Usage_Band_Min__c\" : \"" + unitValue + "\",";
                                                }
                                            }
                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            {
                                                string unitValue = ((Excel.Range)range.Cells[row, 6]).Text;

                                                if (unitValue.IndexOf("kwh", StringComparison.Ordinal) > 0)
                                                {
                                                    unitValue = unitValue.Substring(0, unitValue.IndexOf("kwh", StringComparison.Ordinal));
                                                    if (!unitValue.Equals(string.Empty))
                                                        json += "\"Usage_Band_Max__c\" : \"" + unitValue + "\",";
                                                }
                                            }

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 2]).Text);
                                                if (unitType != string.Empty)
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row, 3]).Text + "\",";
                                            }

                                            for (int innerRow = row; innerRow <= range.Rows.Count; innerRow++)
                                            {
                                                if (methods.GetUniqueIdentifierOE(range, innerRow) == methods.GetUniqueIdentifierOE(range, innerRow + 1))
                                                {
                                                    if (((Excel.Range)range.Cells[innerRow + 1, 2] != null) && (((Excel.Range)range.Cells[innerRow + 1, 2]).Text != string.Empty) && ((Excel.Range)range.Cells[innerRow + 1, 3] != null) && (((Excel.Range)range.Cells[innerRow + 1, 3]).Text != string.Empty))
                                                    {
                                                        unitType = methods.GetUnitTypeFieldName(((Excel.Range)range.Cells[innerRow + 1, 2]).Text);
                                                        if (unitType != string.Empty)
                                                        {
                                                            json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[innerRow + 1, 3]).Text + "\",";
                                                            passToRowNO++;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    break;
                                                }
                                            }


                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            json += "\"Tariff_Type__c\" : \"1\",";


                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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

                                            passToRowNO++;
                                        }
                                    }
                                    else
                                    {
                                        int passToRowNO = 2;
                                        Dictionary<string, string> UnitRateList = new Dictionary<string, string>();
                                        //for (int row = 3; row <= range.Rows.Count; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 4] == null) || (((Excel.Range)range.Cells[row, 4]).Text == string.Empty))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 2]).Text.ToLower() == "standing charge"))
                                            {
                                                if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                                {
                                                    UnitRateList.Add(((Excel.Range)range.Cells[row, 1]).Text, ((Excel.Range)range.Cells[row, 3]).Text);

                                                    passToRowNO++;
                                                    continue;
                                                }
                                            }
                                            else if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 2]).Text.ToLower() == "unit rate"))
                                            {
                                                recordCreated++;
                                                multipleRecordCreateNo++;

                                                json += "{";
                                                json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                                if (UnitRateList.ContainsKey(((Excel.Range)range.Cells[row, 1]).Text))
                                                    json += "    \"Standing_Charge__c\" : \"" + UnitRateList[((Excel.Range)range.Cells[row, 1]).Text] + "\",";
                                                else
                                                    json += "    \"Standing_Charge__c\" : \"0\",";
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 3]).Text + "\",";
                                                if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                                {
                                                    if (supplierNO == 10)
                                                        gasTariffId = methods.GetGasTariffIdOG_REN(((Excel.Range)range.Cells[row, 4]).Text);
                                                    else
                                                        gasTariffId = methods.GetGasTariffIdOG_ACQ(((Excel.Range)range.Cells[row, 4]).Text);

                                                    if (gasTariffId != string.Empty)
                                                        json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                                }
                                                if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                {
                                                    string ldz = ((Excel.Range)range.Cells[row, 5]).Text;
                                                    json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                                }

                                                if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                                {
                                                    if (((Excel.Range)range.Cells[row, 6]).Text == "01")
                                                    {
                                                        json += "    \"Usage_Band_Min__c\" : \"3000\",";
                                                        json += "    \"Usage_Band_Max__c\" : \"73200\",";
                                                    }
                                                    else if (((Excel.Range)range.Cells[row, 6]).Text == "02")
                                                    {
                                                        json += "    \"Usage_Band_Min__c\" : \"73201\",";
                                                        json += "    \"Usage_Band_Max__c\" : \"293000\",";
                                                    }
                                                    else if (((Excel.Range)range.Cells[row, 6]).Text == "03")
                                                    {
                                                        json += "    \"Usage_Band_Min__c\" : \"293001\",";
                                                        json += "    \"Usage_Band_Max__c\" : \"732000\",";
                                                    }
                                                }

                                                json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                                //json += "\"Tariff_Type__c\" : \"1\",";

                                                if (json.Last() == ',')
                                                    json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                                json += "},";

                                                if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                                {
                                                    json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                    json += "]";
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
                                            else
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            passToRowNO++;
                                        }
                                    }
                                    break;
                                }
                            case 12:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        //for (int row = 2; row <= 2; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (((Excel.Range)range.Cells[row, 8] == null) || (((Excel.Range)range.Cells[row, 8]).Text == string.Empty))
                                                continue;

                                            //skip record if Unit Rate is empty
                                            if ((((Excel.Range)range.Cells[row, 25] == null) || (((Excel.Range)range.Cells[row, 25]).Text == string.Empty)))
                                                continue;

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string pesArea = methods.GetPESAreaID(((Excel.Range)range.Cells[row, 3]).Text);
                                                if (pesArea != string.Empty)
                                                    json += "\"PES_Area__c\" : \"" + pesArea + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 4]).Text + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                electricityTariffId = methods.GetElectricityTariffIdSP(((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 9]).Text);
                                                if (electricityTariffId != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 10]).Text), @"[,]", "") + "\",";
                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 11]).Text), @"[,]", "") + "\",";

                                            DateTime earliestDate = DateTime.MinValue;
                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 12]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                earliestDate = DateTime.Parse(date);
                                                //earliestDate = DateTime.ParseExact(date, "MM/dd/yyyy", null);
                                                json += "\"EarliestContractStartDate__c\" : \"" + earliestDate.ToString("yyyy-MM-dd") + "\",";
                                                //earliestContractStartDate = ((Excel.Range)range.Cells[row, 2]).Text;
                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(earliestContractStartDate.Substring(3, 2) + "/" + earliestContractStartDate.Substring(0, 2) + "/" + earliestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 13]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                                //latestContractStartDate = ((Excel.Range)range.Cells[row, 3]).Text;
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(latestContractStartDate.Substring(3, 2) + "/" + latestContractStartDate.Substring(0, 2) + "/" + latestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            else if (earliestDate != DateTime.MinValue)
                                            {
                                                json += "\"LatestContractStartDate__c\" : \"" + earliestDate.AddDays(180).ToString("yyyy-MM-dd") + "\",";
                                                //latestContractStartDate = ((Excel.Range)range.Cells[row, 3]).Text;
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(latestContractStartDate.Substring(3, 2) + "/" + latestContractStartDate.Substring(0, 2) + "/" + latestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 24] != null) && (((Excel.Range)range.Cells[row, 24]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 24]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 25] != null) && (((Excel.Range)range.Cells[row, 25]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 25]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 26] != null) && (((Excel.Range)range.Cells[row, 26]).Text != string.Empty))
                                                json += "    \"Night_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 26]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 27] != null) && (((Excel.Range)range.Cells[row, 27]).Text != string.Empty))
                                                json += "    \"Weekend_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 27]).Text + "\",";

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            json += "\"Tariff_Type__c\" : \"1\",";


                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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
                                    }
                                    else
                                    {
                                        int passToRowNO = 2;
                                        //for (int row = 2; row <= 4; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 8] == null) || (((Excel.Range)range.Cells[row, 8]).Text == string.Empty))
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 16] != null) && (((Excel.Range)range.Cells[row, 16]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 16]).Text != "Monthly Direct Debit"))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 3]).Text;
                                                json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                gasTariffId = methods.GetGasTariffIdSP(((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 9]).Text);
                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 10]).Text), @"[,]", "") + "\",";
                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 11]).Text), @"[,]", "") + "\",";

                                            DateTime earliestDate = DateTime.MinValue;
                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 12]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                earliestDate = DateTime.Parse(date);
                                                //earliestDate = DateTime.ParseExact(date, "MM/dd/yyyy", null);
                                                json += "\"EarliestContractStartDate__c\" : \"" + earliestDate.ToString("yyyy-MM-dd") + "\",";
                                                //earliestContractStartDate = ((Excel.Range)range.Cells[row, 2]).Text;
                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(earliestContractStartDate.Substring(3, 2) + "/" + earliestContractStartDate.Substring(0, 2) + "/" + earliestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 13]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                                //latestContractStartDate = ((Excel.Range)range.Cells[row, 3]).Text;
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(latestContractStartDate.Substring(3, 2) + "/" + latestContractStartDate.Substring(0, 2) + "/" + latestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            else if (earliestDate != DateTime.MinValue)
                                            {
                                                json += "\"LatestContractStartDate__c\" : \"" + earliestDate.AddDays(180).ToString("yyyy-MM-dd") + "\",";
                                                //latestContractStartDate = ((Excel.Range)range.Cells[row, 3]).Text;
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(latestContractStartDate.Substring(3, 2) + "/" + latestContractStartDate.Substring(0, 2) + "/" + latestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 24] != null) && (((Excel.Range)range.Cells[row, 24]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 24]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 28] != null) && (((Excel.Range)range.Cells[row, 28]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 28]).Text + "\",";

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            //json += "\"Tariff_Type__c\" : \"1\",";

                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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

                                            passToRowNO++;
                                        }
                                    }
                                    break;
                                }
                            case 13:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        DateTime firstDate = model.FirstDate;
                                        DateTime secondDate = model.SecondDate;
                                        DateTime thirdDate = model.ThirdDate;
                                        DateTime fourthDate = model.FourthDate;

                                        //for (int row = 3; row <= 4; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (((Excel.Range)range.Cells[row, 8] == null) || (((Excel.Range)range.Cells[row, 8]).Text == string.Empty))
                                                continue;

                                            //skip record if Unit Rate is empty
                                            if (
                                                (((Excel.Range)range.Cells[row, 13] == null) || (((Excel.Range)range.Cells[row, 13]).Text == string.Empty)) &&
                                                (((Excel.Range)range.Cells[row, 14] == null) || (((Excel.Range)range.Cells[row, 14]).Text == string.Empty)) &&
                                                (((Excel.Range)range.Cells[row, 17] == null) || (((Excel.Range)range.Cells[row, 17]).Text == string.Empty))
                                               )
                                                continue;

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string contractDate = ((Excel.Range)range.Cells[row, 3]).Text;
                                                string[] dateArray = contractDate.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    contractDate = "0" + contractDate;
                                                if (dateArray[1].Length == 1)
                                                    contractDate = contractDate.Insert(3, "0");

                                                DateTime date = DateTime.Parse(contractDate);

                                                if (date.Date == firstDate.Date)
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + model.EarliestContractStartDate_First.ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + model.LatestContractStartDate_First.ToString("yyyy-MM-dd") + "\",";
                                                }
                                                else if (date.Date == secondDate.Date)
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + model.EarliestContractStartDate_Second.ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + model.LatestContractStartDate_Second.ToString("yyyy-MM-dd") + "\",";
                                                }
                                                else if (date.Date == thirdDate.Date)
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + model.EarliestContractStartDate_Third.ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + model.LatestContractStartDate_Third.ToString("yyyy-MM-dd") + "\",";
                                                }
                                                else if (date.Date == fourthDate.Date)
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + model.EarliestContractStartDate_Fourth.ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + model.LatestContractStartDate_Fourth.ToString("yyyy-MM-dd") + "\",";
                                                }
                                            }

                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                string pesArea = ((Excel.Range)range.Cells[row, 4]).Text;
                                                pesArea = methods.GetPESAreaID(pesArea.Substring(0, 2));
                                                if (pesArea != string.Empty)
                                                    json += "\"PES_Area__c\" : \"" + pesArea + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            {
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 6]).Text + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                            {
                                                electricityTariffId = methods.GetElectricityTariffIdSSE(((Excel.Range)range.Cells[row, 8]).Text);
                                                if (electricityTariffId != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                                json += "    \"StandingChargeQuarterly__c\" : \"" + ((Excel.Range)range.Cells[row, 11]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                                                json += "    \"StandingChargeQuarterlyAMR__c\" : \"" + ((Excel.Range)range.Cells[row, 12]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 13]).Text + "\",";
                                            else if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 14]).Text + "\",";
                                            else if (((Excel.Range)range.Cells[row, 17] != null) && (((Excel.Range)range.Cells[row, 17]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 17]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 15] != null) && (((Excel.Range)range.Cells[row, 15]).Text != string.Empty))
                                                json += "    \"Weekend_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 15]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 16] != null) && (((Excel.Range)range.Cells[row, 16]).Text != string.Empty))
                                                json += "    \"Night_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 16]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 45] != null) && (((Excel.Range)range.Cells[row, 45]).Text != string.Empty))
                                                json += "    \"FiTCharge__c\" : \"" + ((Excel.Range)range.Cells[row, 45]).Text + "\",";

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            json += "\"Tariff_Type__c\" : \"1\",";


                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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
                                    }
                                    else
                                    {
                                        //for (int row = 3; row <= 4; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (((Excel.Range)range.Cells[row, 3] == null) || (((Excel.Range)range.Cells[row, 3]).Text == string.Empty))
                                                continue;

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                string dateInput = ((Excel.Range)range.Cells[row, 2]).Text;
                                                string[] dateArray = dateInput.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    dateInput = "0" + dateInput;
                                                if (dateArray[1].Length == 1)
                                                    dateInput = dateInput.Insert(3, "0");

                                                DateTime date = DateTime.Parse(dateInput);
                                                //DateTime date = DateTime.ParseExact(dateInput, "MM/dd/yyyy", null);

                                                string monthECSD = date.AddMonths(-1).Month.ToString();
                                                if (monthECSD.Length == 1)
                                                    monthECSD = "0" + monthECSD;

                                                string monthLCSD = date.AddMonths(-1).Month.ToString();
                                                if (monthLCSD.Length == 1)
                                                    monthLCSD = "0" + monthLCSD;


                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(monthECSD + "/15/" + date.AddMonths(-1).Year).ToString("yyyy-MM-dd")) + "\",";
                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(monthECSD + "/15/" + date.AddMonths(-1).Year, "MM/dd/yyyy", null).ToString("yyyy-MM-dd")) + "\",";
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(monthLCSD + "/14/" + date.Year).ToString("yyyy-MM-dd")) + "\",";
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(monthLCSD + "/14/" + date.Year, "MM/dd/yyyy", null).ToString("yyyy-MM-dd")) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                gasTariffId = methods.GetGasTariffIdSSE(((Excel.Range)range.Cells[row, 3]).Text);
                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                                json += "    \"StandingChargeQuarterly__c\" : \"" + ((Excel.Range)range.Cells[row, 6]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 7]).Text + "\",";

                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                            {
                                                string usageBand = ((Excel.Range)range.Cells[row, 8]).Text;
                                                usageBand = usageBand.Replace(" ", string.Empty);

                                                string[] usageBandArray = usageBand.Split('-');
                                                if (usageBandArray.Length == 2)
                                                {
                                                    if (int.TryParse(usageBandArray[0], out int outputMin) && int.TryParse(usageBandArray[1], out int outputMax))
                                                    {
                                                        json += "\"Usage_Band_Min__c\" : \"" + usageBandArray[0] + "\",";
                                                        json += "\"Usage_Band_Max__c\" : \"" + usageBandArray[1] + "\",";
                                                    }
                                                    else
                                                    {
                                                        json += "\"Usage_Band_Min__c\" : \"0\",";
                                                        json += "\"Usage_Band_Max__c\" : \"0\",";
                                                    }
                                                }

                                            }

                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 10]).Text;
                                                json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                            }

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            //json += "\"Tariff_Type__c\" : \"1\",";

                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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
                                    }
                                    break;
                                }
                            case 14:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        ObjectDoesNotExist();
                                        return View("Error");
                                    }
                                    else
                                    {
                                        int passToRowNO = 2;
                                        //for (int row = 2; row <= 4; row++)
                                        for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 12] == null) || (((Excel.Range)range.Cells[row, 12]).Text == string.Empty))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 14]).Text != "DD"))
                                            {
                                                passToRowNO++;
                                                continue;
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 1]).Text), @"[,]", "") + "\",";
                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 2]).Text), @"[,]", "") + "\",";
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 5]).Text;
                                                json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 7]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 8]).Text;
                                                string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                                if (dateArray[0].Length == 1)
                                                    date = "0" + date;
                                                if (dateArray[1].Length == 1)
                                                    date = date.Insert(3, "0");

                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                            {
                                                gasTariffId = methods.GetGasTariffIdCNG(((Excel.Range)range.Cells[row, 12]).Text + ((Excel.Range)range.Cells[row, 13]).Text);
                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 15] != null) && (((Excel.Range)range.Cells[row, 15]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 15]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 16] != null) && (((Excel.Range)range.Cells[row, 16]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 16]).Text + "\",";

                                            json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                            //json += "\"Tariff_Type__c\" : \"1\",";

                                            if (json.Last() == ',')
                                                json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                            json += "},";

                                            if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                            {
                                                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                                json += "]";
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

                                            passToRowNO++;
                                        }
                                    }
                                    break;
                                }
                            default:
                                break;
                        }

                        if (supplierNO == 15)
                        {
                            if (isElectricityTariffPrice)
                            {
                                ObjectDoesNotExist();
                                return View("Error");
                            }
                            else
                            {
                                string crownGasPower = SupplierNO15(ref multipleRecordCreateNo, ref json, ref range, ref methods, ref gasTariffId, ref requestCreate, ref responseCreate, ref uri);
                                if (crownGasPower.Equals("error"))
                                    return View("Error");
                                else
                                    json = crownGasPower;

                                ////for (int row = 3; row <= 4; row++)
                                //for (int row = 2; row <= range.Rows.Count; row++)
                                //{
                                //    if (((Excel.Range)range.Cells[row, 7] == null) || (((Excel.Range)range.Cells[row, 7]).Text == string.Empty))
                                //        continue;

                                //    recordCreated++;
                                //    multipleRecordCreateNo++;

                                //    json += "{";
                                //    json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                //    if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                                //    {
                                //        string ldz = ((Excel.Range)range.Cells[row, 1]).Text;
                                //        json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                //    }
                                //    string standingCharge = string.Empty;
                                //    if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                //    {
                                //        standingCharge = ((Excel.Range)range.Cells[row, 3]).Text;
                                //        json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 3]).Text) * 100) + "\",";
                                //    }
                                //    if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                //        json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 4]).Text + "\",";
                                //    if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                //        json += "    \"Usage_Band_Min__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 5]).Text), @"[,]", "") + "\",";
                                //    if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                //        json += "    \"Usage_Band_Max__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 6]).Text), @"[,]", "") + "\",";
                                //    if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                //    {
                                //        gasTariffId = methods.GetGasTariffIdCG(((Excel.Range)range.Cells[row, 7]).Text, standingCharge);
                                //        if (gasTariffId != string.Empty)
                                //            json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                //    }
                                //    if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                //    {
                                //        string date = ((Excel.Range)range.Cells[row, 8]).Text;
                                //        string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                //        if (dateArray[0].Length == 1)
                                //            date = "0" + date;
                                //        if (dateArray[1].Length == 1)
                                //            date = date.Insert(3, "0");

                                //        json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                //    }
                                //    if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                //    {
                                //        string date = ((Excel.Range)range.Cells[row, 9]).Text;
                                //        string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                //        if (dateArray[0].Length == 1)
                                //            date = "0" + date;
                                //        if (dateArray[1].Length == 1)
                                //            date = date.Insert(3, "0");

                                //        json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                //    }

                                //    json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                //    //json += "\"Tariff_Type__c\" : \"1\",";

                                //    if (json.Last() == ',')
                                //        json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                //    json += "},";

                                //    if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                //    {
                                //        json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                //        json += "]";
                                //        json += "}";

                                //        requestCreate.Content = new StringContent(json, Encoding.UTF8, "application/json");
                                //        responseCreate = Client.SendAsync(requestCreate).Result;
                                //        result = responseCreate.Content.ReadAsStringAsync().Result;

                                //        doc = XDocument.Parse(result);
                                //        if (doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("hasErrors").ElementAt(0).Value.Equals("true"))
                                //        {
                                //            ImportFailed(doc);
                                //            return View("Error");
                                //        }

                                //        requestCreate = new HttpRequestMessage(HttpMethod.Post, uri);
                                //        requestCreate.Headers.Add("Authorization", "Bearer " + AuthToken);
                                //        requestCreate.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
                                //        json = "{";
                                //        json += "\"records\" :[";
                                //        RecordCreated += multipleRecordCreateNo;
                                //        multipleRecordCreateNo = 0;
                                //    }
                                //}
                            }
                        }
                        //else 
                        if ((supplierNO == 16) || (supplierNO == 17))
                        {
                            if (isElectricityTariffPrice)
                            {
                                ObjectDoesNotExist();
                                return View("Error");
                            }
                            else
                            {

                                string dyceEnergy = SupplierNO16_17(ref multipleRecordCreateNo, ref json, ref range, ref methods, ref gasTariffId, ref requestCreate, ref responseCreate, ref uri, ref supplierNO);
                                if (dyceEnergy.Equals("error"))
                                    return View("Error");
                                else
                                    json = dyceEnergy;
                                ////for (int row = 3; row <= 4; row++)
                                //for (int row = 2; row <= range.Rows.Count; row++)
                                //{
                                //    for (int yearRow = 1; yearRow <= 3; yearRow++)
                                //    {
                                //        if ((yearRow == 1) && (((Excel.Range)range.Cells[row, 4] == null) || (((Excel.Range)range.Cells[row, 4]).Text == string.Empty)))
                                //            continue;
                                //        else if ((yearRow == 2) && (((Excel.Range)range.Cells[row, 8] == null) || (((Excel.Range)range.Cells[row, 8]).Text == string.Empty)))
                                //            continue;
                                //        else if ((yearRow == 3) && (((Excel.Range)range.Cells[row, 12] == null) || (((Excel.Range)range.Cells[row, 12]).Text == string.Empty)))
                                //            continue;

                                //        recordCreated++;
                                //        multipleRecordCreateNo++;

                                //        json += "{";
                                //        json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "_" + yearRow + "\"},";

                                //        if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                                //            json += "    \"Usage_Band_Min__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 1]).Text), @"[,]", "") + "\",";
                                //        if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                //            json += "    \"Usage_Band_Max__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 2]).Text), @"[,]", "") + "\",";
                                //        if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                //        {
                                //            string ldz = ((Excel.Range)range.Cells[row, 3]).Text;
                                //            json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                //        }

                                //        if (yearRow == 1)
                                //        {
                                //            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                //            {
                                //                if (supplierNO == 16)
                                //                    gasTariffId = methods.GetGasTariffIdDG_REN(((Excel.Range)range.Cells[row, 4]).Text);
                                //                else
                                //                    gasTariffId = methods.GetGasTariffIdDG_ACQ(((Excel.Range)range.Cells[row, 4]).Text);

                                //                if (gasTariffId != string.Empty)
                                //                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                //            }
                                //            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                //                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";
                                //            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                //                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 6]).Text + "\",";
                                //        }
                                //        else if (yearRow == 2)
                                //        {
                                //            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                //            {
                                //                if (supplierNO == 15)
                                //                    gasTariffId = methods.GetGasTariffIdDG_REN(((Excel.Range)range.Cells[row, 8]).Text);
                                //                else
                                //                    gasTariffId = methods.GetGasTariffIdDG_ACQ(((Excel.Range)range.Cells[row, 8]).Text);

                                //                if (gasTariffId != string.Empty)
                                //                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                //            }
                                //            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                //                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 9]).Text + "\",";
                                //            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                //                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 10]).Text + "\",";
                                //        }
                                //        else if (yearRow == 3)
                                //        {
                                //            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                                //            {
                                //                if (supplierNO == 15)
                                //                    gasTariffId = methods.GetGasTariffIdDG_REN(((Excel.Range)range.Cells[row, 12]).Text);
                                //                else
                                //                    gasTariffId = methods.GetGasTariffIdDG_ACQ(((Excel.Range)range.Cells[row, 12]).Text);

                                //                if (gasTariffId != string.Empty)
                                //                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                //            }
                                //            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                //                json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 13]).Text + "\",";
                                //            if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                                //                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 14]).Text + "\",";
                                //        }

                                //        if (((Excel.Range)range.Cells[row, 16] != null) && (((Excel.Range)range.Cells[row, 16]).Text != string.Empty))
                                //        {
                                //            string date = ((Excel.Range)range.Cells[row, 16]).Text;
                                //            string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                //            if (dateArray[0].Length == 1)
                                //                date = "0" + date;
                                //            if (dateArray[1].Length == 1)
                                //                date = date.Insert(3, "0");

                                //            json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                //        }
                                //        if (((Excel.Range)range.Cells[row, 17] != null) && (((Excel.Range)range.Cells[row, 17]).Text != string.Empty))
                                //        {
                                //            string date = ((Excel.Range)range.Cells[row, 17]).Text;
                                //            string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                                //            if (dateArray[0].Length == 1)
                                //                date = "0" + date;
                                //            if (dateArray[1].Length == 1)
                                //                date = date.Insert(3, "0");

                                //            json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                                //        }

                                //        json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                //        //json += "\"Tariff_Type__c\" : \"1\",";

                                //        if (json.Last() == ',')
                                //            json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                //        json += "},";

                                //        if (yearRow == 3)
                                //        {
                                //            if ((multipleRecordCreateNo == 198) && (row != range.Rows.Count))
                                //            {
                                //                json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                //                json += "]";
                                //                json += "}";

                                //                requestCreate.Content = new StringContent(json, Encoding.UTF8, "application/json");
                                //                responseCreate = Client.SendAsync(requestCreate).Result;
                                //                result = responseCreate.Content.ReadAsStringAsync().Result;

                                //                doc = XDocument.Parse(result);
                                //                if (doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("hasErrors").ElementAt(0).Value.Equals("true"))
                                //                {
                                //                    ImportFailed(doc);
                                //                    return View("Error");
                                //                }

                                //                requestCreate = new HttpRequestMessage(HttpMethod.Post, uri);
                                //                requestCreate.Headers.Add("Authorization", "Bearer " + AuthToken);
                                //                requestCreate.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
                                //                json = "{";
                                //                json += "\"records\" :[";
                                //                RecordCreated += multipleRecordCreateNo;
                                //                multipleRecordCreateNo = 0;
                                //            }
                                //        }
                                //    }
                                //}
                            }

                        }
                        else 
                        if (supplierNO == 18)
                        {
                            if (isElectricityTariffPrice)
                            {
                                ObjectDoesNotExist();
                                return View("Error");
                            }
                            else
                            {
                                //An unhandled exception of type 'System.StackOverflowException' occurred in mscorlib.dll
                                string eon = SupplierNO18(ref multipleRecordCreateNo, ref json, ref range, ref methods, ref gasTariffId, ref requestCreate, ref responseCreate, ref uri);
                                if (eon.Equals("error"))
                                    return View("Error");
                                else
                                    json = eon;


                                //for (int row = 3; row <= 4; row++)
                                ////for (int row = 2; row <= range.Rows.Count; row++)
                                //{
                                //recordCreated++;
                                //multipleRecordCreateNo++;

                                //json += "{";
                                //json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                //if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                //{
                                //    string ldz = ((Excel.Range)range.Cells[row, 2]).Text;
                                //    json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                                //}
                                //if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                //{
                                //    gasTariffId = methods.GetGasTariffIdEON(((Excel.Range)range.Cells[row, 11]).Text, ((Excel.Range)range.Cells[row, 50]).Text + "-" + ((Excel.Range)range.Cells[row, 51]).Text);
                                //    if (gasTariffId != string.Empty)
                                //        json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                //}
                                //if (((Excel.Range)range.Cells[row, 32] != null) && (((Excel.Range)range.Cells[row, 32]).Text != string.Empty))
                                //{
                                //    json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 32]).Text) * 100) + "\",";
                                //}
                                //if (((Excel.Range)range.Cells[row, 36] != null) && (((Excel.Range)range.Cells[row, 36]).Text != string.Empty))
                                //{
                                //    json += "    \"Unit_Rate__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 36]).Text) * 100) + "\",";
                                //}
                                //if (((Excel.Range)range.Cells[row, 50] != null) && (((Excel.Range)range.Cells[row, 50]).Text != string.Empty))
                                //    json += "    \"Usage_Band_Min__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 50]).Text) + "\",";
                                //if (((Excel.Range)range.Cells[row, 51] != null) && (((Excel.Range)range.Cells[row, 51]).Text != string.Empty))
                                //    json += "    \"Usage_Band_Max__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 51]).Text) + "\",";

                                //json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                                ////json += "\"Tariff_Type__c\" : \"1\",";

                                //if (json.Last() == ',')
                                //    json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                                //json += "},";

                                //if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                                //{
                                //    json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                                //    json += "]";
                                //    json += "}";

                                //    requestCreate.Content = new StringContent(json, Encoding.UTF8, "application/json");
                                //    responseCreate = Client.SendAsync(requestCreate).Result;
                                //    result = responseCreate.Content.ReadAsStringAsync().Result;

                                //    doc = XDocument.Parse(result);
                                //    if (doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("hasErrors").ElementAt(0).Value.Equals("true"))
                                //    {
                                //        ImportFailed(doc);
                                //        return View("Error");
                                //    }

                                //    requestCreate = new HttpRequestMessage(HttpMethod.Post, uri);
                                //    requestCreate.Headers.Add("Authorization", "Bearer " + AuthToken);
                                //    requestCreate.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/xml"));
                                //    json = "{";
                                //    json += "\"records\" :[";
                                //    RecordCreated += multipleRecordCreateNo;
                                //    multipleRecordCreateNo = 0;
                                //}
                                //}
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

                        //CloseExcelFile();
                        workBook.Close(true, Missing.Value, Missing.Value);
                        application.Quit();
                        Marshal.ReleaseComObject(workSheet);
                        Marshal.ReleaseComObject(workBook);
                        Marshal.ReleaseComObject(application);

                        Status = "Completed";
                        RecordCreated = recordCreated;
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
                DateTime now = DateTime.Now;
                string logPath = @""+WebConfigurationManager.AppSettings["logPath"]; ;

                if (!System.IO.File.Exists(logPath))
                {
                    System.IO.File.Create(logPath);
                    TextWriter tw = new StreamWriter(logPath);
                    tw.WriteLine("Log - " + now);
                    tw.WriteLine(ex);
                    tw.WriteLine();
                    tw.Close();
                }
                else if (System.IO.File.Exists(logPath))
                {
                    string str;
                    using (StreamReader sreader = new StreamReader(logPath))
                    {
                        str = sreader.ReadToEnd();
                    }

                    System.IO.File.Delete(logPath);

                    using (StreamWriter tw = new StreamWriter(logPath, false))
                    {
                        tw.WriteLine("Log - " + now);
                        tw.WriteLine(ex);
                        tw.WriteLine();
                        tw.Write(str);
                    }
                }

                Status = "Failed";
                RecordFailed = 1;
                MessageError = ex.ToString();

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

                // Create Log record after the import ends
                CreateLog(Client, results);

                throw ex;
            }
        }

        public void ImportFailed(XDocument doc)
        {
            Status = "Failed";
            RecordFailed = 1;
            MessageError = doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("results").ElementAt(0).Descendants("errors").ElementAt(0).Descendants("message").ElementAt(0).Value;
            StatusCode = doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("results").ElementAt(0).Descendants("errors").ElementAt(0).Descendants("statusCode").ElementAt(0).Value;
            ReferenceId = doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("results").ElementAt(0).Descendants("referenceId").ElementAt(0).Value;

            //CloseExcelFile();
            workBook.Close(true, null, null);
            application.Quit();
            Marshal.ReleaseComObject(workSheet);
            Marshal.ReleaseComObject(workBook);
            Marshal.ReleaseComObject(application);

            //PopulateOutputTable();
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

            // Create Log record after the import ends
            CreateLog(Client, results);
        }

        public void ObjectDoesNotExist()
        {
            Status = "Failed";
            RecordFailed = 0;
            MessageError = "It can't import this file for the selected object. This supplier is not connected with " + Object + ".";
            StatusCode = "Failed";
            ReferenceId = "first row of Excel file";

            //CloseExcelFile();
            workBook.Close(true, null, null);
            application.Quit();
            Marshal.ReleaseComObject(workSheet);
            Marshal.ReleaseComObject(workBook);
            Marshal.ReleaseComObject(application);

            //PopulateOutputTable();
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

            // Create Log record after the import ends
            CreateLog(Client, results);
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

            // Create Log record after the import ends
            CreateLog(Client, results);
        }

        public void CloseExcelFile()
        {
            workBook.Close(true, null, null);
            application.Quit();
            Marshal.ReleaseComObject(workSheet);
            Marshal.ReleaseComObject(workBook);
            Marshal.ReleaseComObject(application);
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

        private void CreateLog(HttpClient client, Results results)
        {
            string createMessage = $"<root>" +
                $"<Name>{"ImportDataFromExcel - "+results.StartDate.ToString()}</Name>" +
                $"<Status__c>{results.Status}</Status__c>" +
                $"<Object__c>{results.Object}</Object__c>" +
                $"<Record_Created__c>{results.RecordCreated}</Record_Created__c>" +
                $"<Record_Failed__c>{results.RecordFailed}</Record_Failed__c>" +
                $"<Start_Date__c>{results.StartDate.ToString()}</Start_Date__c>" +
                $"<Processing_Time__c>{results.ProcessingTime}</Processing_Time__c>" +
                $"<Message_Error__c>{results.MessageError}</Message_Error__c>" +
                $"<Status_Code__c>{results.StatusCode}</Status_Code__c>" +
                $"<Reference_Id__c>{results.ReferenceId}</Reference_Id__c>" +
                $"</root>";

            string result = CreateRecord(client, createMessage, "Log__c");
        }

        public string SupplierNO15(ref int multipleRecordCreateNo, ref string json, ref Excel.Range range, ref Methods methods, ref string gasTariffId, ref HttpRequestMessage requestCreate, ref HttpResponseMessage responseCreate, ref string uri)
        {
            //for (int row = 3; row <= 4; row++)
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                if (((Excel.Range)range.Cells[row, 7] == null) || (((Excel.Range)range.Cells[row, 7]).Text == string.Empty))
                    continue;

                recordCreated++;
                multipleRecordCreateNo++;

                json += "{";
                json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                {
                    string ldz = ((Excel.Range)range.Cells[row, 1]).Text;
                    json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                }
                string standingCharge = string.Empty;
                if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                {
                    standingCharge = ((Excel.Range)range.Cells[row, 3]).Text;
                    json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 3]).Text) * 100) + "\",";
                }
                if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                    json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 4]).Text + "\",";
                if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                    json += "    \"Usage_Band_Min__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 5]).Text), @"[,]", "") + "\",";
                if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                    json += "    \"Usage_Band_Max__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 6]).Text), @"[,]", "") + "\",";
                if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                {
                    gasTariffId = methods.GetGasTariffIdCG(((Excel.Range)range.Cells[row, 7]).Text, standingCharge);
                    if (gasTariffId != string.Empty)
                        json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                }
                if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                {
                    string date = ((Excel.Range)range.Cells[row, 8]).Text;
                    string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                    if (dateArray[0].Length == 1)
                        date = "0" + date;
                    if (dateArray[1].Length == 1)
                        date = date.Insert(3, "0");

                    json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                    //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                }
                if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                {
                    string date = ((Excel.Range)range.Cells[row, 9]).Text;
                    string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                    if (dateArray[0].Length == 1)
                        date = "0" + date;
                    if (dateArray[1].Length == 1)
                        date = date.Insert(3, "0");

                    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                    //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.ParseExact(date, "MM/dd/yyyy", null)).ToString("yyyy-MM-dd") + "\",";
                }

                json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
                //json += "\"Tariff_Type__c\" : \"1\",";

                if (json.Last() == ',')
                    json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                json += "},";

                if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                {
                    json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                    json += "]";
                    json += "}";

                    requestCreate.Content = new StringContent(json, Encoding.UTF8, "application/json");
                    responseCreate = Client.SendAsync(requestCreate).Result;
                    string result = responseCreate.Content.ReadAsStringAsync().Result;

                    XDocument doc = XDocument.Parse(result);
                    if (doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("hasErrors").ElementAt(0).Value.Equals("true"))
                    {
                        ImportFailed(doc);
                        //return View("Error");
                        return "error";
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

            return json;
        }

        public string SupplierNO16_17(ref int multipleRecordCreateNo, ref string json, ref Excel.Range range, ref Methods methods, ref string gasTariffId, ref HttpRequestMessage requestCreate, ref HttpResponseMessage responseCreate, ref string uri, ref int supplierNO)
        {
            //for (int row = 3; row <= 4; row++)
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                for (int yearRow = 1; yearRow <= 3; yearRow++)
                {
                    if ((yearRow == 1) && (((Excel.Range)range.Cells[row, 4] == null) || (((Excel.Range)range.Cells[row, 4]).Text == string.Empty)))
                        continue;
                    else if ((yearRow == 2) && (((Excel.Range)range.Cells[row, 8] == null) || (((Excel.Range)range.Cells[row, 8]).Text == string.Empty)))
                        continue;
                    else if ((yearRow == 3) && (((Excel.Range)range.Cells[row, 12] == null) || (((Excel.Range)range.Cells[row, 12]).Text == string.Empty)))
                        continue;

                    recordCreated++;
                    multipleRecordCreateNo++;

                    json += "{";
                    json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "_" + yearRow + "\"},";

                    if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                        json += "    \"Usage_Band_Min__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 1]).Text), @"[,]", "") + "\",";
                    if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                        json += "    \"Usage_Band_Max__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 2]).Text), @"[,]", "") + "\",";
                    if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                    {
                        string ldz = ((Excel.Range)range.Cells[row, 3]).Text;
                        json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                    }

                    if (yearRow == 1)
                    {
                        if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                        {
                            if (supplierNO == 16)
                                gasTariffId = methods.GetGasTariffIdDG_REN(((Excel.Range)range.Cells[row, 4]).Text);
                            else
                                gasTariffId = methods.GetGasTariffIdDG_ACQ(((Excel.Range)range.Cells[row, 4]).Text);

                            if (gasTariffId != string.Empty)
                                json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                        }
                        if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                            json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";
                        if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                            json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 6]).Text + "\",";
                    }
                    else if (yearRow == 2)
                    {
                        if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                        {
                            if (supplierNO == 15)
                                gasTariffId = methods.GetGasTariffIdDG_REN(((Excel.Range)range.Cells[row, 8]).Text);
                            else
                                gasTariffId = methods.GetGasTariffIdDG_ACQ(((Excel.Range)range.Cells[row, 8]).Text);

                            if (gasTariffId != string.Empty)
                                json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                        }
                        if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                            json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 9]).Text + "\",";
                        if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                            json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 10]).Text + "\",";
                    }
                    else if (yearRow == 3)
                    {
                        if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                        {
                            if (supplierNO == 15)
                                gasTariffId = methods.GetGasTariffIdDG_REN(((Excel.Range)range.Cells[row, 12]).Text);
                            else
                                gasTariffId = methods.GetGasTariffIdDG_ACQ(((Excel.Range)range.Cells[row, 12]).Text);

                            if (gasTariffId != string.Empty)
                                json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                        }
                        if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                            json += "    \"Standing_Charge__c\" : \"" + ((Excel.Range)range.Cells[row, 13]).Text + "\",";
                        if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                            json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 14]).Text + "\",";
                    }

                    if (((Excel.Range)range.Cells[row, 16] != null) && (((Excel.Range)range.Cells[row, 16]).Text != string.Empty))
                    {
                        string date = ((Excel.Range)range.Cells[row, 16]).Text;
                        string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                        if (dateArray[0].Length == 1)
                            date = "0" + date;
                        if (dateArray[1].Length == 1)
                            date = date.Insert(3, "0");

                        //string logPath = @"C:/Users/Renis Kraja/Desktop/Salesforce/ImportDataFromExcel/Logs/Log.txt";

                        //if (!System.IO.File.Exists(logPath))
                        //{
                        //    System.IO.File.Create(logPath);
                        //    TextWriter tw = new StreamWriter(logPath);
                        //    tw.WriteLine("date - " + date);
                        //    tw.WriteLine();
                        //    tw.Close();
                        //}
                        //else if (System.IO.File.Exists(logPath))
                        //{
                        //    string str;
                        //    using (StreamReader sreader = new StreamReader(logPath))
                        //    {
                        //        str = sreader.ReadToEnd();
                        //    }

                        //    System.IO.File.Delete(logPath);

                        //    using (StreamWriter tw = new StreamWriter(logPath, false))
                        //    {
                        //        tw.WriteLine("date - " + date);
                        //        tw.WriteLine();
                        //        tw.Write(str);
                        //    }
                        //}


                        DateTime parsedDate = DateTime.Parse(date);
                        //DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", null);
                        json += "\"EarliestContractStartDate__c\" : \"" + parsedDate.ToString("yyyy-MM-dd") + "\",";
                    }
                    if (((Excel.Range)range.Cells[row, 17] != null) && (((Excel.Range)range.Cells[row, 17]).Text != string.Empty))
                    {
                        string date = ((Excel.Range)range.Cells[row, 17]).Text;
                        string[] dateArray = date.Split(new string[] { "/" }, 3, StringSplitOptions.None);
                        if (dateArray[0].Length == 1)
                            date = "0" + date;
                        if (dateArray[1].Length == 1)
                            date = date.Insert(3, "0");

                        //DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", null);
                        DateTime parsedDate = DateTime.Parse(date);
                        json += "\"LatestContractStartDate__c\" : \"" + (parsedDate).ToString("yyyy-MM-dd") + "\",";
                    }

                    DateTime now = DateTime.Now;
                    json += "\"Pricing_Start__c\" : \"" + now.ToString("yyyy-MM-dd") + "\",";
                    //json += "\"Tariff_Type__c\" : \"1\",";

                    if (json.Last() == ',')
                        json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                    json += "},";

                    if (yearRow == 3)
                    {
                        if ((multipleRecordCreateNo == 198) && (row != range.Rows.Count))
                        {
                            json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                            json += "]";
                            json += "}";

                            requestCreate.Content = new StringContent(json, Encoding.UTF8, "application/json");
                            responseCreate = Client.SendAsync(requestCreate).Result;
                            string result = responseCreate.Content.ReadAsStringAsync().Result;

                            XDocument doc = XDocument.Parse(result);
                            if (doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("hasErrors").ElementAt(0).Value.Equals("true"))
                            {
                                ImportFailed(doc);
                                //return View("Error");
                                return "error";
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
                }
            }

            return json;
        }

        public string SupplierNO18(ref int multipleRecordCreateNo, ref string json, ref Excel.Range range, ref Methods methods, ref string gasTariffId, ref HttpRequestMessage requestCreate, ref HttpResponseMessage responseCreate, ref string uri)
        {
            //for (int row = 2; row <= 10; row++)
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                if (((Excel.Range)range.Cells[row, 11] == null) || (((Excel.Range)range.Cells[row, 11]).Text == string.Empty))
                    continue;

                if (((Excel.Range)range.Cells[row, 36] == null) || (((Excel.Range)range.Cells[row, 36]).Text == string.Empty))
                    continue;
                else if (!double.TryParse(Regex.Replace((((Excel.Range)range.Cells[row, 36]).Text), @"[,]", ""), out double parsedNumber))
                    continue;

                recordCreated++;
                multipleRecordCreateNo++;

                json += "{";
                json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                {
                    string ldz = ((Excel.Range)range.Cells[row, 2]).Text;
                    json += "\"PES_Area__c\" : \"" + methods.GetLDZ_ID(ldz) + "\",";
                }
                if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                {
                    gasTariffId = methods.GetGasTariffIdEON(((Excel.Range)range.Cells[row, 11]).Text, ((Excel.Range)range.Cells[row, 50]).Text + "-" + ((Excel.Range)range.Cells[row, 51]).Text);
                    if (gasTariffId != string.Empty)
                        json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                }
                if (((Excel.Range)range.Cells[row, 32] != null) && (((Excel.Range)range.Cells[row, 32]).Text != string.Empty))
                {
                    if (double.TryParse(Regex.Replace((((Excel.Range)range.Cells[row, 32]).Text), @"[,]", ""), out double parsedNumber))
                        json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(Regex.Replace((((Excel.Range)range.Cells[row, 32]).Text), @"[,]", "")) * 100) + "\",";
                }
                if (((Excel.Range)range.Cells[row, 36] != null) && (((Excel.Range)range.Cells[row, 36]).Text != string.Empty))
                {
                    json += "    \"Unit_Rate__c\" : \"" + (Convert.ToDouble(Regex.Replace((((Excel.Range)range.Cells[row, 36]).Text), @"[,]", "")) * 100) + "\",";
                }
                if (((Excel.Range)range.Cells[row, 50] != null) && (((Excel.Range)range.Cells[row, 50]).Text != string.Empty))
                    json += "    \"Usage_Band_Min__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 50]).Text), @"[,]", "") + "\",";
                if (((Excel.Range)range.Cells[row, 51] != null) && (((Excel.Range)range.Cells[row, 51]).Text != string.Empty))
                    json += "    \"Usage_Band_Max__c\" : \"" + Regex.Replace((((Excel.Range)range.Cells[row, 51]).Text), @"[,]", "") + "\",";

                json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";

                if (json.Last() == ',')
                    json = json.Remove(json.Length - 1, 1); // Remove the last "," if the last cell of the Excel file is empty

                json += "},";

                if ((multipleRecordCreateNo == 200) && (row != range.Rows.Count))
                {
                    json = json.Remove(json.Length - 1, 1); // Remove "," from the last record added to the json
                    json += "]";
                    json += "}";

                    requestCreate.Content = new StringContent(json, Encoding.UTF8, "application/json");
                    responseCreate = Client.SendAsync(requestCreate).Result;
                    string result = responseCreate.Content.ReadAsStringAsync().Result;

                    XDocument doc = XDocument.Parse(result);
                    if (doc.Descendants("SObjectTreeResponse").ElementAt(0).Descendants("hasErrors").ElementAt(0).Value.Equals("true"))
                    {
                        ImportFailed(doc);
                        //return View("Error");
                        return "error";
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

            return json;
        }
    }
}