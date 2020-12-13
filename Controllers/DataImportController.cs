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

        private SelectList suppliers = new SelectList(new[]
        {
            new { ID = "1", Name = "British Gas Lite" },
            new { ID = "2", Name = "British Gas" },
            new { ID = "3", Name = "British Gas DSC" },
            new { ID = "4", Name = "Smartest Energy Electric" },
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
                        int recordCreated = 0;
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
                                        int passToRowNO = 3;
                                        for (int row = 2; row <= 8; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

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
                                                    day = day.Remove(day.Length - 1, 1);

                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1] + "/" + day + "/" + date.Substring(date.Length - 4))).ToString("yyyy-MM-dd") + "\",";
                                                
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 2]).Text;

                                                string day = date.Substring(0, 2);
                                                if (day.Last() == '-')
                                                    day = day.Remove(day.Length - 1, 1);

                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1] + "/" + day + "/" + date.Substring(date.Length - 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                                json += "\"PES_Area__c\" : \"" + GetPESAreaID(((Excel.Range)range.Cells[row, 4]).Text) + "\",";
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                electricityTariffId = GetElectricityTariffIdBGL(((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 9]).Text);
                                                if (electricityTariffId != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariffId + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                            {
                                                if (int.TryParse(((Excel.Range)range.Cells[row, 10]).Text, out int output))
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"" + GetUsageBandMin(Int32.Parse(((Excel.Range)range.Cells[row, 10]).Text)) + "\",";
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
                                                unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 13]).Text);
                                                if (unitType != string.Empty)
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row, 14]).Text + "\",";
                                            }

                                            for (int innerRow = row; innerRow <= range.Rows.Count; innerRow++)
                                            {
                                                if (GetUniqueIdentifierBGL(range, innerRow) == GetUniqueIdentifierBGL(range, innerRow + 1))
                                                {
                                                    if (((Excel.Range)range.Cells[innerRow + 1, 13] != null) && (((Excel.Range)range.Cells[innerRow + 1, 13]).Text != string.Empty) && ((Excel.Range)range.Cells[innerRow + 1, 14] != null) && (((Excel.Range)range.Cells[innerRow + 1, 14]).Text != string.Empty))
                                                    {
                                                        unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[innerRow + 1, 13]).Text);
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
                                        int passToRowNO = 3;
                                        for (int row = 2; row <= 6; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

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
                                                if (GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) != string.Empty)
                                                {
                                                    string day = date.Substring(0, 2);
                                                    if (day.Last() == '-')
                                                        day = day.Remove(day.Length - 1, 1);
                                                    string year = date.Substring(date.Length - 2);
                                                    if (year == "20")
                                                        year = "2020";
                                                    else if (year == "21")
                                                        year = "2021";

                                                    json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + year)).ToString("yyyy-MM-dd") + "\",";
                                                }
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 3]).Text;
                                                if (GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) != string.Empty)
                                                {
                                                    string day = date.Substring(0, 2);
                                                    if (day.Last() == '-')
                                                        day = day.Remove(day.Length - 1, 1);
                                                    string year = date.Substring(date.Length - 2);
                                                    if (year == "20")
                                                        year = "2020";
                                                    else if (year == "21")
                                                        year = "2021";

                                                    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + year)).ToString("yyyy-MM-dd") + "\",";
                                                }
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 4]).Text;
                                                json += "    \"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                            {
                                                electricityTariffId = GetGasTariffIdBGL(((Excel.Range)range.Cells[row, 7]).Text + ((Excel.Range)range.Cells[row, 8]).Text);
                                                if (electricityTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + electricityTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                if (int.TryParse(((Excel.Range)range.Cells[row, 9]).Text, out int output))
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"" + GetUsageBandMinGas(Int32.Parse(((Excel.Range)range.Cells[row, 9]).Text)) + "\",";
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
                                                unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 12]).Text);
                                                if (unitType != string.Empty)
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row, 13]).Text + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row + 1, 12] != null) && (((Excel.Range)range.Cells[row + 1, 12]).Text != string.Empty) && ((Excel.Range)range.Cells[row + 1, 13] != null) && (((Excel.Range)range.Cells[row + 1, 13]).Text != string.Empty))
                                            {
                                                unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[row + 1, 12]).Text);
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
                                        int passToRowNO = 3;
                                        for (int row = 2; row <= 498; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 13]).Text != "DD"))
                                            {
                                                string test = ((Excel.Range)range.Cells[row, 13]).Text;
                                                passToRowNO++;
                                                continue;
                                            }

                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(((Excel.Range)range.Cells[row, 2]).Text)).ToString("yyyy-MM-dd") + "\",";
                                                //earliestContractStartDate = ((Excel.Range)range.Cells[row, 2]).Text;
                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(earliestContractStartDate.Substring(3, 2) + "/" + earliestContractStartDate.Substring(0, 2) + "/" + earliestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(((Excel.Range)range.Cells[row, 3]).Text)).ToString("yyyy-MM-dd") + "\",";
                                                //latestContractStartDate = ((Excel.Range)range.Cells[row, 3]).Text;
                                                //json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(latestContractStartDate.Substring(3, 2) + "/" + latestContractStartDate.Substring(0, 2) + "/" + latestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                                json += "\"PES_Area__c\" : \"" + GetPESAreaID(((Excel.Range)range.Cells[row, 4]).Text) + "\",";
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                electricityTariffId = GetElectricityTariffIdBGL(((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 9]).Text);
                                                if (electricityTariffId != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariffId + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                            {
                                                if (int.TryParse(((Excel.Range)range.Cells[row, 11]).Text, out int output))
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"" + GetUsageBandMin(Int32.Parse(((Excel.Range)range.Cells[row, 11]).Text)) + "\",";
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
                                                unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 14]).Text);
                                                if (unitType != string.Empty)
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row, 15]).Text + "\",";
                                            }

                                            for (int innerRow = row; innerRow <= range.Rows.Count; innerRow++)
                                            {
                                                if (GetUniqueIdentifierBG(range, innerRow) == GetUniqueIdentifierBG(range, innerRow + 1))
                                                {
                                                    if (((Excel.Range)range.Cells[innerRow + 1, 14] != null) && (((Excel.Range)range.Cells[innerRow + 1, 14]).Text != string.Empty) && ((Excel.Range)range.Cells[innerRow + 1, 15] != null) && (((Excel.Range)range.Cells[innerRow + 1, 15]).Text != string.Empty))
                                                    {
                                                        unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[innerRow + 1, 14]).Text);
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
                                        int passToRowNO = 3;
                                        for (int row = 2; row <= 6; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

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
                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 3]).Text;
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 4]).Text;
                                                json += "    \"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                            {
                                                if (supplierNO == 2)
                                                    gasTariffId = GetGasTariffIdBG(((Excel.Range)range.Cells[row, 7]).Text + ((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 10]).Text);
                                                else
                                                    gasTariffId = GetGasTariffIdBG_DSC(((Excel.Range)range.Cells[row, 7]).Text + ((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 10]).Text);

                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                if (int.TryParse(((Excel.Range)range.Cells[row, 9]).Text, out int output))
                                                {
                                                    json += "\"Usage_Band_Min__c\" : \"" + GetUsageBandMinGas(Int32.Parse(((Excel.Range)range.Cells[row, 9]).Text)) + "\",";
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
                                                unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 12]).Text);
                                                if (unitType != string.Empty)
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row, 13]).Text + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row + 1, 12] != null) && (((Excel.Range)range.Cells[row + 1, 12]).Text != string.Empty) && ((Excel.Range)range.Cells[row + 1, 13] != null) && (((Excel.Range)range.Cells[row + 1, 13]).Text != string.Empty))
                                            {
                                                unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[row + 1, 12]).Text);
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
                                        for (int row = 3; row <= 3; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {

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
                                                json += "\"PES_Area__c\" : \"" + GetPESAreaID(((Excel.Range)range.Cells[row, 2]).Text) + "\",";
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";

                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            {
                                                string electricityTariff = GetElectricityTariffIdSE(((Excel.Range)range.Cells[row, 6]).Text);
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
                                                json += "    \"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(((Excel.Range)range.Cells[row, 13]).Text)).ToString("yyyy-MM-dd") + "\",";
                                            if (((Excel.Range)range.Cells[row, 14] != null) && (((Excel.Range)range.Cells[row, 14]).Text != string.Empty))
                                                json += "    \"LatestContractStartDate__c\" : \"" + (DateTime.Parse(((Excel.Range)range.Cells[row, 14]).Text)).ToString("yyyy-MM-dd") + "\",";

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
                                        for (int row = 4; row <= 5; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                gasTariffId = GetGasTariffIdSE(((Excel.Range)range.Cells[row, 3]).Text + ((Excel.Range)range.Cells[row, 2]).Text);
                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 4]).Text;
                                                json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
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
                                        int passToRowNO = 3;
                                        for (int row = 2; row <= 7; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty) && (((Excel.Range)range.Cells[row, 2]).Text == "HH"))
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
                                                    string pesArea = GetPESAreaID(((Excel.Range)range.Cells[row, 1]).Text);
                                                    if (pesArea != string.Empty)
                                                        json += "\"PES_Area__c\" : \"" + pesArea + "\",";
                                                }
                                                if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                                    json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 2]).Text + "\",";

                                                if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                {
                                                    unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 5]).Text);
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
                                                    if (GetUniqueIdentifierVE(range, innerRow) == GetUniqueIdentifierVE(range, innerRow + 1))
                                                    {
                                                        unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[innerRow + 1, 5]).Text);
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
                                        for (int row = 3; row <= 61; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {

                                            if (
                                                ((Excel.Range)range.Cells[row, 3] != null)
                                                &&
                                                (((Excel.Range)range.Cells[row, 3]).Text != string.Empty)
                                                &&
                                                ((((Excel.Range)range.Cells[row, 3]).Text == "OUN"))
                                               )
                                            {
                                                continue;
                                            }

                                            if (
                                                ((Excel.Range)range.Cells[row, 11] != null)
                                                &&
                                                (((Excel.Range)range.Cells[row, 11]).Text != string.Empty)
                                                &&
                                                (!((((Excel.Range)range.Cells[row, 11]).Text == "0") || (((Excel.Range)range.Cells[row, 11]).Text != "0.0")))
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
                                                json += "\"PES_Area__c\" : \"" + GetPESAreaID(region.Substring(region.Length - 2)) + "\",";
                                            }


                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string profileClass = GetProfileClassEDF(((Excel.Range)range.Cells[row, 3]).Text);
                                                if (profileClass != string.Empty)
                                                    json += "\"Profile_Code__c\" : \"" + profileClass + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                string electricityTariff = GetElectricityTariffIdEDF(((Excel.Range)range.Cells[row, 5]).Text);
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
                                        int passToRowNO = 3;
                                        for (int row = 2; row <= 4; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

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
                                                json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                gasTariffId = GetGasTariffIdEDF(((Excel.Range)range.Cells[row, 5]).Text);
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
                                        for (int row = 3; row <= 4; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
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
                                                string electricityTariff = GetElectricityTariffIdGazprom(((Excel.Range)range.Cells[row, 4]).Text);
                                                if (electricityTariff != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariff + "\",";
                                            }
                                            
                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            {
                                                string pesArea = GetPESAreaID(((Excel.Range)range.Cells[row, 6]).Text);
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
                                                    if (GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) != string.Empty)
                                                    {
                                                        string day = date.Substring(0, 2);
                                                        if (day.Last() == '-')
                                                            day = day.Remove(day.Length - 1, 1);

                                                        json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + date.Substring(date.Length - 4))).ToString("yyyy-MM-dd") + "\",";
                                                    }
                                                }
                                                else if (date.ToLower().Contains("after"))
                                                {
                                                    date = date.Substring(date.Length - 11); //31-Jan-2021
                                                    if (GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) != string.Empty)
                                                    {
                                                        string day = date.Substring(0, 2);
                                                        if (day.Last() == '-')
                                                            day = day.Remove(day.Length - 1, 1);

                                                        json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + date.Substring(date.Length - 4))).ToString("yyyy-MM-dd") + "\",";
                                                        json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(GetMonth(date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + day + "/" + date.Substring(date.Length - 4)).AddMonths(6)).ToString("yyyy-MM-dd") + "\",";
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
                                        for (int row = 4; row <= 5; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            {
                                                if (supplierNO == 7)
                                                    gasTariffId = GetGasTariffIdGP_REN(((Excel.Range)range.Cells[row, 2]).Text + ((Excel.Range)range.Cells[row, 6]).Text);
                                                else
                                                    gasTariffId = GetGasTariffIdGP_ACQ(((Excel.Range)range.Cells[row, 2]).Text + ((Excel.Range)range.Cells[row, 6]).Text);


                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + ((Excel.Range)range.Cells[row, 4]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 5]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 7]).Text;
                                                json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 8]).Text;
                                                if (date.Contains("-"))
                                                    json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(GetMonth(date.Substring(date.Length - 2) + "/" + date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + date.Substring(0, 4))).ToString("yyyy-MM-dd") + "\",";
                                                else
                                                    json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 9]).Text;
                                                if (date.Contains("-"))
                                                    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(GetMonth(date.Substring(date.Length - 2) + "/" + date.Split(new string[] { "-" }, 3, StringSplitOptions.None)[1]) + "/" + date.Substring(0, 4))).ToString("yyyy-MM-dd") + "\",";
                                                else
                                                    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                                json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 10]).Text) * 100) + "\",";
                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                                json += "    \"Unit_Rate__c\" : \"" + ((Excel.Range)range.Cells[row, 11]).Text + "\",";

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
                                        for (int row = 3; row <= 4; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                string pesArea = GetPESAreaID(((Excel.Range)range.Cells[row, 2]).Text);
                                                if (pesArea != string.Empty)
                                                    json += "\"PES_Area__c\" : \"" + pesArea + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 3]).Text + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                string month = GetMonth(((Excel.Range)range.Cells[row, 4]).Text);
                                                if (month != string.Empty)
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(month + "/1/" + ((Excel.Range)range.Cells[row, 5]).Text)).ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(month + "/" + DateTime.DaysInMonth(Convert.ToInt32(((Excel.Range)range.Cells[row, 5]).Text), Convert.ToInt32(month)) + "/" + ((Excel.Range)range.Cells[row, 5]).Text).ToString("yyyy-MM-dd")) + "\",";
                                                }
                                            }
                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            {
                                                string electricityTariff = GetElectricityTariffIdNpower(((Excel.Range)range.Cells[row, 6]).Text);
                                                if (electricityTariff != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariff + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + ((Excel.Range)range.Cells[row, 9]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + ((Excel.Range)range.Cells[row, 10]).Text+ "\",";
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
                                        for (int row = 3; row <= 4; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 2]).Text;
                                                json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                string month = GetMonth(((Excel.Range)range.Cells[row, 4]).Text);
                                                if (month != string.Empty)
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(month + "/1/" + ((Excel.Range)range.Cells[row, 5]).Text)).ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(month + "/" + DateTime.DaysInMonth(Convert.ToInt32(((Excel.Range)range.Cells[row, 5]).Text), Convert.ToInt32(month)) + "/" + ((Excel.Range)range.Cells[row, 5]).Text).ToString("yyyy-MM-dd")) + "\",";
                                                }
                                            }
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + ((Excel.Range)range.Cells[row, 8]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + Decimal.Floor(Convert.ToDecimal(((Excel.Range)range.Cells[row, 9]).Text)) + "\",";
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                            {
                                                gasTariffId = GetGasTariffIdNpower(((Excel.Range)range.Cells[row, 10]).Text);
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
                                        int passToRowNO = 3;
                                        for (int row = 2; row <= 7; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

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
                                                json += "\"PES_Area__c\" : \"" + GetPESAreaID(profileClass.Substring(0, 2)) + "\",";
                                                json += "\"Profile_Code__c\" : \"" + profileClass.Substring(2, 1) + "\",";

                                                if (supplierNO == 10)
                                                    electricityTariffId = GetElectricityTariffIdOE_REN(profileClass);
                                                else
                                                    electricityTariffId = GetElectricityTariffIdOE_ACQ(profileClass);

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
                                                unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[row, 2]).Text);
                                                if (unitType != string.Empty)
                                                    json += "\"" + unitType + "\" : \"" + ((Excel.Range)range.Cells[row, 3]).Text + "\",";
                                            }

                                            for (int innerRow = row; innerRow <= range.Rows.Count; innerRow++)
                                            {
                                                if (GetUniqueIdentifierOE(range, innerRow) == GetUniqueIdentifierOE(range, innerRow + 1))
                                                {
                                                    if (((Excel.Range)range.Cells[innerRow + 1, 2] != null) && (((Excel.Range)range.Cells[innerRow + 1, 2]).Text != string.Empty) && ((Excel.Range)range.Cells[innerRow + 1, 3] != null) && (((Excel.Range)range.Cells[innerRow + 1, 3]).Text != string.Empty))
                                                    {
                                                        unitType = GetUnitTypeFieldName(((Excel.Range)range.Cells[innerRow + 1, 2]).Text);
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
                                        int passToRowNO = 3;
                                        Dictionary<string, string> UnitRateList = new Dictionary<string, string>();
                                        for (int row = 3; row <= range.Rows.Count; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

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
                                                        gasTariffId = GetGasTariffIdOG_REN(((Excel.Range)range.Cells[row, 4]).Text);
                                                    else
                                                        gasTariffId = GetGasTariffIdOG_ACQ(((Excel.Range)range.Cells[row, 4]).Text);

                                                    if (gasTariffId != string.Empty)
                                                        json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                                }
                                                if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                                {
                                                    string ldz = ((Excel.Range)range.Cells[row, 5]).Text;
                                                    json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
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
                                                continue;

                                            passToRowNO++;
                                        }
                                    }
                                    break;
                                }
                            case 12:
                                {
                                    if (isElectricityTariffPrice)
                                    {
                                        for (int row = 3; row <= 4; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string pesArea = GetPESAreaID(((Excel.Range)range.Cells[row, 3]).Text);
                                                if (pesArea != string.Empty)
                                                    json += "\"PES_Area__c\" : \"" + pesArea + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 4]).Text + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                electricityTariffId = GetElectricityTariffIdSP(((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 9]).Text);
                                                if (electricityTariffId != string.Empty)
                                                    json += "\"Electricity_Tariff__c\" : \"" + electricityTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + ((Excel.Range)range.Cells[row, 10]).Text + "\",";
                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 11]).Text) + "\",";

                                            DateTime earliestDate = DateTime.MinValue;
                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                                            {
                                                earliestDate = DateTime.Parse(((Excel.Range)range.Cells[row, 12]).Text);
                                                json += "\"EarliestContractStartDate__c\" : \"" + earliestDate.ToString("yyyy-MM-dd") + "\",";
                                                //earliestContractStartDate = ((Excel.Range)range.Cells[row, 2]).Text;
                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(earliestContractStartDate.Substring(3, 2) + "/" + earliestContractStartDate.Substring(0, 2) + "/" + earliestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                            {
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(((Excel.Range)range.Cells[row, 13]).Text)).ToString("yyyy-MM-dd") + "\",";
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
                                        int passToRowNO = 3;
                                        for (int row = 2; row <= 4; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
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
                                                json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
                                            }

                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                            {
                                                gasTariffId = GetGasTariffIdSP(((Excel.Range)range.Cells[row, 8]).Text + ((Excel.Range)range.Cells[row, 9]).Text);
                                                if (gasTariffId != string.Empty)
                                                    json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                                                json += "    \"Usage_Band_Min__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 10]).Text) + "\",";
                                            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 11]).Text) + "\",";

                                            DateTime earliestDate = DateTime.MinValue;
                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                                            {
                                                earliestDate = DateTime.Parse(((Excel.Range)range.Cells[row, 12]).Text);
                                                json += "\"EarliestContractStartDate__c\" : \"" + earliestDate.ToString("yyyy-MM-dd") + "\",";
                                                //earliestContractStartDate = ((Excel.Range)range.Cells[row, 2]).Text;
                                                //json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(earliestContractStartDate.Substring(3, 2) + "/" + earliestContractStartDate.Substring(0, 2) + "/" + earliestContractStartDate.Substring(6, 4))).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                            {
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(((Excel.Range)range.Cells[row, 13]).Text)).ToString("yyyy-MM-dd") + "\",";
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
                                        for (int row = 3; row <= 4; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                string contractDate = ((Excel.Range)range.Cells[row, 3]).Text;
                                                if (((contractDate.Equals("12/1/2020"))) || (contractDate.Equals("12/01/2020")) || (contractDate.Equals("01/12/2020")))
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + model.EarliestContractStartDate_First.ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + model.LatestContractStartDate_First.ToString("yyyy-MM-dd") + "\",";
                                                }
                                                else if ((contractDate.Equals("4/1/2021")) || (contractDate.Equals("04/01/2021")) || (contractDate.Equals("01/04/2021")))
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + model.EarliestContractStartDate_Second.ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + model.LatestContractStartDate_Second.ToString("yyyy-MM-dd") + "\",";
                                                }
                                                else if ((contractDate.Equals("10/1/2021")) || (contractDate.Equals("10/01/2021")) || (contractDate.Equals("01/10/2021")))
                                                {
                                                    json += "\"EarliestContractStartDate__c\" : \"" + model.EarliestContractStartDate_Third.ToString("yyyy-MM-dd") + "\",";
                                                    json += "\"LatestContractStartDate__c\" : \"" + model.LatestContractStartDate_Third.ToString("yyyy-MM-dd") + "\",";
                                                }
                                            }

                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                string pesArea = ((Excel.Range)range.Cells[row, 4]).Text;
                                                pesArea = GetPESAreaID(pesArea.Substring(0, 2));
                                                if (pesArea != string.Empty)
                                                    json += "\"PES_Area__c\" : \"" + pesArea + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                            {
                                                json += "\"Profile_Code__c\" : \"" + ((Excel.Range)range.Cells[row, 6]).Text + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                            {
                                                electricityTariffId = GetElectricityTariffIdSSE(((Excel.Range)range.Cells[row, 8]).Text);
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
                                        for (int row = 3; row <= 4; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            recordCreated++;
                                            multipleRecordCreateNo++;

                                            json += "{";
                                            json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            {
                                                DateTime date = DateTime.Parse(((Excel.Range)range.Cells[row, 2]).Text);
                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date.AddMonths(-1).Month + "/15/" + date.AddMonths(-1).Year).ToString("yyyy-MM-dd")) + "\",";
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date.Month + "/14/" + date.Year).ToString("yyyy-MM-dd")) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                            {
                                                gasTariffId = GetGasTariffIdSSE(((Excel.Range)range.Cells[row, 3]).Text);
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
                                                json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
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
                                        int passToRowNO = 3;
                                        for (int row = 2; row <= 4; row++)
                                        //for (int row = 2; row <= range.Rows.Count; row++)
                                        {
                                            if (row != passToRowNO)
                                                continue;

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
                                                json += "    \"Usage_Band_Min__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 1]).Text) + "\",";
                                            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                                json += "    \"Usage_Band_Max__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 2]).Text) + "\",";
                                            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                                            {
                                                string ldz = ((Excel.Range)range.Cells[row, 5]).Text;
                                                json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 7]).Text;
                                                json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                            {
                                                string date = ((Excel.Range)range.Cells[row, 8]).Text;
                                                json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                            }
                                            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty) && ((Excel.Range)range.Cells[row, 13] != null) && (((Excel.Range)range.Cells[row, 13]).Text != string.Empty))
                                            {
                                                gasTariffId = GetGasTariffIdCNG(((Excel.Range)range.Cells[row, 12]).Text + ((Excel.Range)range.Cells[row, 13]).Text);
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
                                for (int row = 3; row <= 4; row++)
                                //for (int row = 2; row <= range.Rows.Count; row++)
                                {
                                    recordCreated++;
                                    multipleRecordCreateNo++;

                                    json += "{";
                                    json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                    if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                                    {
                                        string ldz = ((Excel.Range)range.Cells[row, 1]).Text;
                                        json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
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
                                        json += "    \"Usage_Band_Min__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 5]).Text) + "\",";
                                    if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                                        json += "    \"Usage_Band_Max__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 6]).Text) + "\",";
                                    if (((Excel.Range)range.Cells[row, 7] != null) && (((Excel.Range)range.Cells[row, 7]).Text != string.Empty))
                                    {
                                        gasTariffId = GetGasTariffIdCG(((Excel.Range)range.Cells[row, 7]).Text, standingCharge);
                                        if (gasTariffId != string.Empty)
                                            json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                    }
                                    if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                                    {
                                        string date = ((Excel.Range)range.Cells[row, 8]).Text;
                                        json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                    }
                                    if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                                    {
                                        string date = ((Excel.Range)range.Cells[row, 9]).Text;
                                        json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
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
                                for (int row = 3; row <= 4; row++)
                                //for (int row = 2; row <= range.Rows.Count; row++)
                                {
                                    for (int yearRow = 1; yearRow <= 3; yearRow++)
                                    {
                                        recordCreated++;
                                        multipleRecordCreateNo++;

                                        json += "{";
                                        json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                        if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                                            json += "    \"Usage_Band_Min__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 1]).Text) + "\",";
                                        if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                            json += "    \"Usage_Band_Max__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 2]).Text) + "\",";
                                        if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                                        {
                                            string ldz = ((Excel.Range)range.Cells[row, 3]).Text;
                                            json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
                                        }

                                        if (yearRow == 1)
                                        {
                                            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                                            {
                                                if (supplierNO == 16)
                                                    gasTariffId = GetGasTariffIdDG_REN(((Excel.Range)range.Cells[row, 4]).Text);
                                                else
                                                    gasTariffId = GetGasTariffIdDG_ACQ(((Excel.Range)range.Cells[row, 4]).Text);

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
                                                    gasTariffId = GetGasTariffIdDG_REN(((Excel.Range)range.Cells[row, 8]).Text);
                                                else
                                                    gasTariffId = GetGasTariffIdDG_ACQ(((Excel.Range)range.Cells[row, 8]).Text);

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
                                                    gasTariffId = GetGasTariffIdDG_REN(((Excel.Range)range.Cells[row, 12]).Text);
                                                else
                                                    gasTariffId = GetGasTariffIdDG_ACQ(((Excel.Range)range.Cells[row, 12]).Text);

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
                                            json += "\"EarliestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                        }
                                        if (((Excel.Range)range.Cells[row, 17] != null) && (((Excel.Range)range.Cells[row, 17]).Text != string.Empty))
                                        {
                                            string date = ((Excel.Range)range.Cells[row, 17]).Text;
                                            json += "\"LatestContractStartDate__c\" : \"" + (DateTime.Parse(date)).ToString("yyyy-MM-dd") + "\",";
                                        }

                                        json += "\"Pricing_Start__c\" : \"" + DateTime.Now.ToString("yyyy-MM-dd") + "\",";
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
                                }
                            }

                        }
                        else if (supplierNO == 18)
                        {
                            if (isElectricityTariffPrice)
                            {
                                ObjectDoesNotExist();
                                return View("Error");
                            }
                            else
                            {
                                for (int row = 3; row <= 4; row++)
                                //for (int row = 2; row <= range.Rows.Count; row++)
                                {
                                    recordCreated++;
                                    multipleRecordCreateNo++;

                                    json += "{";
                                    json += "\"attributes\" : {\"type\" : \"" + Object + "\", \"referenceId\" : \"Row " + row + "\"},";

                                    if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                                    {
                                        string ldz = ((Excel.Range)range.Cells[row, 2]).Text;
                                        json += "\"PES_Area__c\" : \"" + GetLDZ_ID(ldz) + "\",";
                                    }
                                    if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                                    {
                                        gasTariffId = GetGasTariffIdEON(((Excel.Range)range.Cells[row, 11]).Text, ((Excel.Range)range.Cells[row, 50]).Text + "-" + ((Excel.Range)range.Cells[row, 51]).Text);
                                        if (gasTariffId != string.Empty)
                                            json += "\"Gas_Tariff__c\" : \"" + gasTariffId + "\",";
                                    }
                                    if (((Excel.Range)range.Cells[row, 32] != null) && (((Excel.Range)range.Cells[row, 32]).Text != string.Empty))
                                    {
                                        json += "    \"Standing_Charge__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 32]).Text) * 100) + "\",";
                                    }
                                    if (((Excel.Range)range.Cells[row, 36] != null) && (((Excel.Range)range.Cells[row, 36]).Text != string.Empty))
                                    {
                                        json += "    \"Unit_Rate__c\" : \"" + (Convert.ToDouble(((Excel.Range)range.Cells[row, 36]).Text) * 100) + "\",";
                                    }
                                    if (((Excel.Range)range.Cells[row, 50] != null) && (((Excel.Range)range.Cells[row, 50]).Text != string.Empty))
                                        json += "    \"Usage_Band_Min__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 50]).Text) + "\",";
                                    if (((Excel.Range)range.Cells[row, 51] != null) && (((Excel.Range)range.Cells[row, 51]).Text != string.Empty))
                                        json += "    \"Usage_Band_Max__c\" : \"" + Convert.ToDouble(((Excel.Range)range.Cells[row, 51]).Text) + "\",";

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
                //DateTime now = DateTime.Now;
                //string logPath = @"C:/Users/Renis Kraja/Desktop/Salesforce/ImportDataFromExcel/Logs/Log.txt";

                //if (!System.IO.File.Exists(logPath))
                //{
                //    System.IO.File.Create(logPath);
                //    TextWriter tw = new StreamWriter(logPath);
                //    tw.WriteLine("Log - " + now);
                //    tw.WriteLine(ex);
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
                //        tw.WriteLine("Log - " + now);
                //        tw.WriteLine(ex);
                //        tw.WriteLine();
                //        tw.Write(str);
                //    }
                //}

                throw ex;
            }
        }

        public string GetUnitTypeFieldName(string cellValue)
        {
            string unitTypeFieldName = string.Empty;

            switch (cellValue.ToLower())
            {
                case "day unit charge":
                    unitTypeFieldName = "Unit_Rate__c";
                    break;
                case "day rate":
                    unitTypeFieldName = "Unit_Rate__c";
                    break;
                case "night unit charge":
                    unitTypeFieldName = "Night_Rate__c";
                    break;
                case "night rate":
                    unitTypeFieldName = "Night_Rate__c";
                    break;
                case "weekday day unit charge":
                    unitTypeFieldName = "Unit_Rate__c";
                    break;
                case "weekday rate":
                    unitTypeFieldName = "Unit_Rate__c";
                    break;
                case "evening & weekend unit charge":
                    unitTypeFieldName = "Weekend_Rate__c";
                    break;
                case "eve, weekend & night rate":
                    unitTypeFieldName = "Weekend_Rate__c";
                    break;
                case "standing charge":
                    unitTypeFieldName = "Standing_Charge__c";
                    break;
                case "unit charge":
                    unitTypeFieldName = "Unit_Rate__c";
                    break;
                case "unit rate":
                    unitTypeFieldName = "Unit_Rate__c";
                    break;
                default:
                    unitTypeFieldName = string.Empty;
                    break;
            }

            return unitTypeFieldName;
        }

        public string GetPESAreaID(string cellValue)
        {
            string PES = string.Empty;

            switch (cellValue)
            {
                case "10":
                    PES = "a0Z30000004Rl9D";
                    break;
                case "11":
                    PES = "a0Z30000004RlDQ";
                    break;
                case "12":
                    PES = "a0Z30000004RlDV";
                    break;
                case "13":
                    PES = "a0Z30000004RlDa";
                    break;
                case "14":
                    PES = "a0Z30000004RlDf";
                    break;
                case "15":
                    PES = "a0Z30000004RlDk";
                    break;
                case "16":
                    PES = "a0Z30000004RlDp";
                    break;
                case "17":
                    PES = "a0Z30000004RlEO";
                    break;
                case "18":
                    PES = "a0Z30000004RlEJ";
                    break;
                case "19":
                    PES = "a0Z30000004RlDu";
                    break;
                case "20":
                    PES = "a0Z30000004RlDz";
                    break;
                case "21":
                    PES = "a0Z30000004RlE9";
                    break;
                case "22":
                    PES = "a0Z30000004RlE4";
                    break;
                case "23":
                    PES = "a0Z30000004RlEE";
                    break;
                default:
                    PES = string.Empty;
                    break;
            }

            return PES;
        }

        public string GetLDZ_ID(string cellValue)
        {
            string LDZ = string.Empty;

            switch (cellValue)
            {
                case "EA":
                    LDZ = "a0Z30000004Rl9D";
                    break;
                case "EA1":
                    LDZ = "a0Z13000012M0ob";
                    break;
                case "EA2":
                    LDZ = "a0Z13000012M0oc";
                    break;
                case "EA3":
                    LDZ = "a0Z13000012M0od";
                    break;
                case "EA4":
                    LDZ = "a0Z13000012M0oe";
                    break;
                case "EM":
                    LDZ = "a0Z30000004RlDQ";
                    break;
                case "EM1":
                    LDZ = "a0Z13000012M0of";
                    break;
                case "EM2":
                    LDZ = "a0Z13000012M0og";
                    break;
                case "EM3":
                    LDZ = "a0Z13000012M0oh";
                    break;
                case "EM4":
                    LDZ = "a0Z13000012M0oi";
                    break;
                case "LC":
                    LDZ = "a0Z13000012M0oj";
                    break;
                case "LO":
                    LDZ = "a0Z13000012M0ok";
                    break;
                case "LS":
                    LDZ = "a0Z13000012M0ol";
                    break;
                case "LT":
                    LDZ = "a0Z13000012M0om";
                    break;
                case "LW":
                    LDZ = "a0Z13000012M0on";
                    break;
                case "NE":
                    LDZ = "a0Z30000004RlEE";
                    break;
                case "NE1":
                    LDZ = "a0Z13000012M0oo";
                    break;
                case "NE2":
                    LDZ = "a0Z13000012M0op";
                    break;
                case "NE3":
                    LDZ = "a0Z13000012M0oq";
                    break;
                case "NO":
                    LDZ = "a0Z30000004RlDk";
                    break;
                case "NO1":
                    LDZ = "a0Z13000012M0or";
                    break;
                case "NO2":
                    LDZ = "a0Z13000012M0os";
                    break;
                case "NT":
                    LDZ = "a0Z30000004RlDV";
                    break;
                case "NT1":
                    LDZ = "a0Z13000012M0ot";
                    break;
                case "NT2":
                    LDZ = "a0Z13000012M0ou";
                    break;
                case "NT3":
                    LDZ = "a0Z13000012M0ov";
                    break;
                case "NW":
                    LDZ = "a0Z30000004RlDp";
                    break;
                case "NW1":
                    LDZ = "a0Z13000012M0ow";
                    break;
                case "NW2":
                    LDZ = "a0Z13000012M0ox";
                    break;
                case "SC":
                    LDZ = "a0Z30000004RlEJ";
                    break;
                case "SC1":
                    LDZ = "a0Z13000012M0oy";
                    break;
                case "SC2":
                    LDZ = "a0Z13000012M0oz";
                    break;
                case "SC4":
                    LDZ = "a0Z13000012M0p0";
                    break;
                case "SE":
                    LDZ = "a0Z30000004RlDu";
                    break;
                case "SE1":
                    LDZ = "a0Z13000012M0p1";
                    break;
                case "SE2":
                    LDZ = "a0Z13000012M0p2";
                    break;
                case "SO":
                    LDZ = "a0Z30000004RlDz";
                    break;
                case "SO1":
                    LDZ = "a0Z13000012M0p3";
                    break;
                case "SO2":
                    LDZ = "a0Z13000012M0p4";
                    break;
                case "SW":
                    LDZ = "a0Z30000004RlE4";
                    break;
                case "SW1":
                    LDZ = "a0Z13000012M0p5";
                    break;
                case "SW2":
                    LDZ = "a0Z13000012M0p6";
                    break;
                case "SW3":
                    LDZ = "a0Z13000012M0p7";
                    break;
                case "WA":
                    LDZ = "a0Z13000012M0ox";
                    break;
                case "WA1":
                    LDZ = "a0Z13000012M0p8";
                    break;
                case "WA2":
                    LDZ = "a0Z13000012M0p9";
                    break;
                case "WM":
                    LDZ = "a0Z30000004RlDf";
                    break;
                case "WM1":
                    LDZ = "a0Z13000012M0pA";
                    break;
                case "WM2":
                    LDZ = "a0Z13000012M0pB";
                    break;
                case "WM3":
                    LDZ = "a0Z13000012M0pC";
                    break;
                case "WN":
                    LDZ = "a0Z30000004RlDa";
                    break;
                case "WS":
                    LDZ = "a0Z30000004RlE9";
                    break;
                default:
                    LDZ = string.Empty;
                    break;
            }

            return LDZ;
        }

        public string GetMonth(string cellValue)
        {
            string month = string.Empty;

            switch (cellValue.ToLower())
            {
                case "jan":
                    month = "01";
                    break;
                case "feb":
                    month = "02";
                    break;
                case "mar":
                    month = "03";
                    break;
                case "apr":
                    month = "04";
                    break;
                case "may":
                    month = "05";
                    break;
                case "jun":
                    month = "06";
                    break;
                case "jul":
                    month = "07";
                    break;
                case "aug":
                    month = "08";
                    break;
                case "sep":
                    month = "09";
                    break;
                case "oct":
                    month = "10";
                    break;
                case "nov":
                    month = "11";
                    break;
                case "dec":
                    month = "12";
                    break;
                default:
                    month = string.Empty;
                    break;
            }

            return month;
        }

        public string GetElectricityTariffIdBGL(string cellValue)
        {
            string electricityTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "acquisition12":
                    electricityTariffId = "a0h1B00000FLP7H";
                    break;
                case "acquisition24":
                    electricityTariffId = "a0h1B00000FLP7M";
                    break;
                case "acquisition36":
                    electricityTariffId = "a0h1B00000FLP7R";
                    break;
                case "acquisition48":
                    electricityTariffId = "a0h1B00000FLP7W";
                    break;
                case "acquisition60":
                    electricityTariffId = "a0h1B00000FLP7b";
                    break;
                case "renewal12":
                    electricityTariffId = "a0h1B00000Y6pt2";
                    break;
                case "renewal24":
                    electricityTariffId = "a0h1B00000Y6pt7";
                    break;
                case "renewal36":
                    electricityTariffId = "a0h1B00000Y6ptC";
                    break;
                case "renewal48":
                    electricityTariffId = "a0h1B00000Y6ptH";
                    break;
                case "renewal60":
                    electricityTariffId = "a0h1B00000Y6ptM";
                    break;
                default:
                    electricityTariffId = string.Empty;
                    break;
            }

            return electricityTariffId;
        }

        public string GetElectricityTariffIdEDF(string cellValue)
        {
            string electricityTariffId = string.Empty;

            switch (cellValue)
            {
                case "12":
                    electricityTariffId = "a0h1300000P0Vmq";
                    break;
                case "24":
                    electricityTariffId = "a0h1300000P0Vmv";
                    break;
                case "36":
                    electricityTariffId = "a0h1300000P0Vn0";
                    break;
                case "48":
                    electricityTariffId = "a0h1B00000WsNjr";
                    break;
                default:
                    electricityTariffId = string.Empty;
                    break;
            }

            return electricityTariffId;
        }

        public string GetElectricityTariffIdSE(string cellValue)
        {
            string electricityTariffId = string.Empty;

            switch (cellValue)
            {
                case "SmartFIX – 1 Year Renewal":
                    electricityTariffId = "a0h1300000P0Vmq";
                    break;
                case "SmartFIX – 2 Year Renewal":
                    electricityTariffId = "a0h4v00000YAHBn";
                    break;
                case "SmartFIX – 3 Year Renewal":
                    electricityTariffId = "a0h4v00000YAHBs";
                    break;
                case "SmartFIX – 5 Year Renewal":
                    electricityTariffId = "a0h4v00000YAHBx";
                    break;
                case "SmartTRACKER Renewal":
                    electricityTariffId = "a0h4v00000YAHCb";
                    break;
                case "SmartPAY12_Renewal":
                    electricityTariffId = "a0h4v00000YAHCM";
                    break;
                case "SmartPAY24_Renewal":
                    electricityTariffId = "a0h4v00000YAHCR";
                    break;
                case "SmartPAY36_Renewal":
                    electricityTariffId = "a0h4v00000YAHCW";
                    break;
                case "SmartFIX – 1 Year":
                    electricityTariffId = "a0h4v00000YAHBO";
                    break;
                case "SmartFIX – 2 Year":
                    electricityTariffId = "a0h4v00000YAHBT";
                    break;
                case "SmartFIX – 3 Year":
                    electricityTariffId = "a0h4v00000YAHBY";
                    break;
                case "SmartFIX – 5 Year":
                    electricityTariffId = "a0h4v00000YAHBd";
                    break;
                case "SmartTRACKER":
                    electricityTariffId = "a0h4v00000YAHC2";
                    break;
                case "SmartPAY12":
                    electricityTariffId = "a0h4v00000YAHC7";
                    break;
                case "SmartPAY24":
                    electricityTariffId = "a0h4v00000YAHCC";
                    break;
                case "SmartPAY36":
                    electricityTariffId = "a0h4v00000YAHCH";
                    break;
                default:
                    electricityTariffId = string.Empty;
                    break;
            }

            return electricityTariffId;
        }

        public string GetElectricityTariffIdGazprom(string cellValue)
        {
            string electricityTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "1 year":
                    electricityTariffId = "a0h1300000P0Z8m";
                    break;
                case "2 year":
                    electricityTariffId = "a0h1300000P0Z8r";
                    break;
                case "3 year":
                    electricityTariffId = "a0h1300000P0Z8w";
                    break;
                default:
                    electricityTariffId = string.Empty;
                    break;
            }

            return electricityTariffId;
        }

        public string GetElectricityTariffIdNpower(string cellValue)
        {
            string electricityTariffId = string.Empty;

            switch (cellValue)
            {
                case "1":
                    electricityTariffId = "a0h1B00000Uhgyn";
                    break;
                case "2":
                    electricityTariffId = "a0h1B00000Uhgys";
                    break;
                case "3":
                    electricityTariffId = "a0h1B00000Uhgyx";
                    break;
                default:
                    electricityTariffId = string.Empty;
                    break;
            }

            return electricityTariffId;
        }

        public string GetElectricityTariffIdOE_REN(string cellValue)
        {
            string electricityTariffId = string.Empty;

            int dotLocation = cellValue.IndexOf(".", StringComparison.Ordinal);

            if (dotLocation > 0)
            {
                electricityTariffId = cellValue.Substring(0, dotLocation);

                if (electricityTariffId.Substring(electricityTariffId.Length - 2).ToLower().Equals("st"))
                    return "a0ha000000N9Rrm";

                if (electricityTariffId.Substring(electricityTariffId.Length - 3).ToLower().Equals("st4"))
                    return "a0h1300000TlJx0";

                switch (electricityTariffId.Substring(electricityTariffId.Length - 4))
                {
                    case "ren2":
                        electricityTariffId = "a0h1300000UE5pY";
                        break;
                    case "ren3":
                        electricityTariffId = "a0h1300000UE5pd";
                        break;
                    default:
                        electricityTariffId = string.Empty;
                        break;
                }

            }

            return electricityTariffId;
        }

        public string GetElectricityTariffIdOE_ACQ(string cellValue)
        {
            string electricityTariffId = string.Empty;

            int dotLocation = cellValue.IndexOf(".", StringComparison.Ordinal);

            if (dotLocation > 0)
            {
                electricityTariffId = cellValue.Substring(0, dotLocation);

                if (electricityTariffId.Substring(electricityTariffId.Length - 2).ToLower().Equals("st"))
                    return "a0h1300000P0VOF";

                switch (electricityTariffId.Substring(electricityTariffId.Length - 3))
                {
                    case "st2":
                        electricityTariffId = "a0ha000000C6kHd";
                        break;
                    case "st3":
                        electricityTariffId = "a0ha000000C6kHi";
                        break;
                    case "st4":
                        electricityTariffId = "a0h1300000TlJx5";
                        break;
                    default:
                        electricityTariffId = string.Empty;
                        break;
                }

            }

            return electricityTariffId;
        }

        public string GetElectricityTariffIdSP(string cellValue)
        {
            string electricityTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "acquisition12":
                    electricityTariffId = "a0h3000000AdWdC";
                    break;
                case "acquisition24":
                    electricityTariffId = "a0h3000000AdWdH";
                    break;
                case "acquisition36":
                    electricityTariffId = "a0h3000000AdWdM";
                    break;
                case "renewal12":
                    electricityTariffId = "a0h1300000VEBts";
                    break;
                case "renewal24":
                    electricityTariffId = "a0h1300000VEBtx";
                    break;
                case "renewal36":
                    electricityTariffId = "'a0h1300000VEBu2'";
                    break;
                default:
                    electricityTariffId = string.Empty;
                    break;
            }

            return electricityTariffId;
        }

        public string GetElectricityTariffIdSSE(string cellValue)
        {
            string electricityTariffId = string.Empty;

            switch (cellValue)
            {
                case "12":
                    electricityTariffId = "a0h1300000OzIuo";
                    break;
                case "24":
                    electricityTariffId = "a0h1300000OzIut";
                    break;
                case "36":
                    electricityTariffId = "a0h1300000OzIuy";
                    break;
                case "48":
                    electricityTariffId = "a0h1300000OzIv3";
                    break;
                case "60":
                    electricityTariffId = "a0h1B00000WscUR";
                    break;
                default:
                    electricityTariffId = string.Empty;
                    break;
            }

            return electricityTariffId;
        }

        public string GetGasTariffIdBGL(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "acquisition12":
                    gasTariffId = "a0b1B00000PHWi7";
                    break;
                case "acquisition24":
                    gasTariffId = "a0b1B00000PHWiC";
                    break;
                case "acquisition36":
                    gasTariffId = "a0b1B00000PHWiH";
                    break;
                case "acquisition48":
                    gasTariffId = "a0b1B00000FDNWa";
                    break;
                case "acquisition60":
                    gasTariffId = "a0b1B00000FDinL";
                    break;
                case "renewal12":
                    gasTariffId = "a0b1B00000QgEQD";
                    break;
                case "renewal24":
                    gasTariffId = "a0b1B00000QgEQI";
                    break;
                case "renewal36":
                    gasTariffId = "a0b1B00000QgEQX";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdBG(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "renewal12sc":
                    gasTariffId = "a0b1300000NnDHr";
                    break;
                case "renewal24sc":
                    gasTariffId = "a0b1300000NnDI1";
                    break;
                case "renewal36sc":
                    gasTariffId = "a0b1300000NnDIB";
                    break;
                case "renewal48sc":
                    gasTariffId = "a0b1B00000PTLna";
                    break;
                case "renewal60sc":
                    gasTariffId = "a0b1B00000PTLnk";
                    break;
                case "renewal12nsc":
                    gasTariffId = "a0b1300000NnDHw";
                    break;
                case "renewal24nsc":
                    gasTariffId = "a0b1300000NnDI6";
                    break;
                case "renewal36nsc":
                    gasTariffId = "a0b1300000NnDIG";
                    break;
                case "renewal48nsc":
                    gasTariffId = "a0b1B00000PTLnf";
                    break;
                case "renewal60nsc":
                    gasTariffId = "a0b1B00000PTLnp";
                    break;
                case "acquisition12sc":
                    gasTariffId = "a0b30000002iZIJ";
                    break;
                case "acquisition24sc":
                    gasTariffId = "a0b30000002iZIO";
                    break;
                case "acquisition36sc":
                    gasTariffId = "a0b30000002iZIT";
                    break;
                case "acquisition48sc":
                    gasTariffId = "a0b1B00000PTLnG";
                    break;
                case "acquisition60sc":
                    gasTariffId = "a0b1B00000PTLnQ";
                    break;
                case "acquisition12nsc":
                    gasTariffId = "a0b30000002iZIY";
                    break;
                case "acquisition24nsc":
                    gasTariffId = "a0b30000002iZId";
                    break;
                case "acquisition36nsc":
                    gasTariffId = "a0b30000002jf6H";
                    break;
                case "acquisition48nsc":
                    gasTariffId = "a0b1B00000PTLnL";
                    break;
                case "acquisition60nsc":
                    gasTariffId = "a0b1B00000PTLnV";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdBG_DSC(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "acquisition12sc":
                    gasTariffId = "a0b1B00000QhWsi";
                    break;
                case "acquisition24sc":
                    gasTariffId = "a0b1B00000QhWsn";
                    break;
                case "acquisition36sc":
                    gasTariffId = "a0b1B00000QhWss";
                    break;
                case "acquisition48sc":
                    gasTariffId = "a0b1B00000QhWsY";
                    break;
                case "acquisition60sc":
                    gasTariffId = "a0b1B00000QhWsd";
                    break;
                case "renewal12sc":
                    gasTariffId = "a0b1B00000QhWsx";
                    break;
                case "renewal24sc":
                    gasTariffId = "a0b1B00000QhWt2";
                    break;
                case "renewal36sc":
                    gasTariffId = "a0b1B00000QhWt7";
                    break;
                case "renewal48sc":
                    gasTariffId = "a0b1B00000QhWtH";
                    break;
                case "renewal60sc":
                    gasTariffId = "a0b1B00000QhWtM";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdEDF(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue)
            {
                case "12":
                    gasTariffId = "a0b1300000NlFZ1";
                    break;
                case "24":
                    gasTariffId = "a0b1300000NlFZB";
                    break;
                case "36":
                    gasTariffId = "a0b1300000NlFZV";
                    break;
                case "48":
                    gasTariffId = "a0b1B00000Q13fS";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdGP_REN(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "high12 months":
                    gasTariffId = "a0b1B00000Pa5GR";
                    break;
                case "high24 months":
                    gasTariffId = "a0b1B00000Pa5GS";
                    break;
                case "high36 months":
                    gasTariffId = "a0b1B00000Pa5GT";
                    break;
                case "high48 months":
                    gasTariffId = "a0b1B00000Pa5GK";
                    break;
                case "high60 months":
                    gasTariffId = "a0b1B00000Pa5GL";
                    break;
                case "low12 months":
                    gasTariffId = "a0b1B00000Pa5GM";
                    break;
                case "low24 months":
                    gasTariffId = "a0b1B00000Pa5GN";
                    break;
                case "low36 months":
                    gasTariffId = "a0b1B00000Pa5GO";
                    break;
                case "low48 months":
                    gasTariffId = "a0b1B00000Pa5GP";
                    break;
                case "low60 months":
                    gasTariffId = "a0b1B00000Pa5GQ";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdGP_ACQ(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "high12 months":
                    gasTariffId = "a0ba000000C4cgG";
                    break;
                case "high24 months":
                    gasTariffId = "a0ba000000C4cga";
                    break;
                case "high36 months":
                    gasTariffId = "a0ba000000C4cgf";
                    break;
                case "high48 months":
                    gasTariffId = "a0b1300000JTJdn";
                    break;
                case "high60 months":
                    gasTariffId = "a0b1300000JTJds";
                    break;
                case "low12 months":
                    gasTariffId = "a0b1300000LzSKx";
                    break;
                case "low24 months":
                    gasTariffId = "a0b1300000LzSL7";
                    break;
                case "low36 months":
                    gasTariffId = "a0b1300000LzSLC";
                    break;
                case "low48 months":
                    gasTariffId = "a0b1300000LzSmm";
                    break;
                case "low60 months":
                    gasTariffId = "a0b1300000LzSmn";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdNpower(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue)
            {
                case "1":
                    gasTariffId = "a0b1300000Lvhpg";
                    break;
                case "2":
                    gasTariffId = "a0b1300000Lvhpl";
                    break;
                case "3":
                    gasTariffId = "a0b1300000Lvhpq";
                    break;
                case "4":
                    gasTariffId = "a0b1B00000PSgLC";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdOG_REN(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue)
            {
                case "12 Months Existing Business Acquisition & Retention Gas Tariff no S/C":
                    gasTariffId = "a0b1300000No4Xi";
                    break;
                case "12 Months Existing Business Acquisition & Retention Gas Tariff":
                    gasTariffId = "a0b1300000No4XO";
                    break;
                case "24 Months Existing Business Acquisition & Retention Gas Tariff no S/C":
                    gasTariffId = "a0b1300000No4Xn";
                    break;
                case "24 Months Existing Business Acquisition & Retention Gas Tariff":
                    gasTariffId = "a0b1300000No4XY";
                    break;
                case "36 Months Existing Business Acquisition & Retention Gas Tariff no S/C":
                    gasTariffId = "a0b1300000No4Xs";
                    break;
                case "36 Months Existing Business Acquisition & Retention Gas Tariff":
                    gasTariffId = "a0b1300000No4Xd";
                    break;
                case "48 Months Existing Business Acquisition & Retention Gas Tariff no S/C":
                    gasTariffId = "a0b1B00000P0lPh";
                    break;
                case "48 Months Existing Business Acquisition & Retention Gas Tariff":
                    gasTariffId = "a0b1B00000P0lPc";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdSP(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "acquisition12":
                    gasTariffId = "a0b30000002iZJ7";
                    break;
                case "acquisition24":
                    gasTariffId = "a0b30000002iZJC";
                    break;
                case "acquisition36":
                    gasTariffId = "a0b1300000MGHXm";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdOG_ACQ(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue)
            {
                case "12 Months Existing Business Acquisition & Retention Gas Tariff no S/C":
                    gasTariffId = "a0ba000000HVYoc";
                    break;
                case "12 Months Existing Business Acquisition & Retention Gas Tariff":
                    gasTariffId = "a0ba000000HVYny";
                    break;
                case "24 Months Existing Business Acquisition & Retention Gas Tariff no S/C":
                    gasTariffId = "a0ba000000Gwy1G";
                    break;
                case "24 Months Existing Business Acquisition & Retention Gas Tariff":
                    gasTariffId = "a0ba000000GxO8n";
                    break;
                case "36 Months Existing Business Acquisition & Retention Gas Tariff no S/C":
                    gasTariffId = "a0ba000000GxNts";
                    break;
                case "36 Months Existing Business Acquisition & Retention Gas Tariff":
                    gasTariffId = "a0ba000000GxOFB";
                    break;
                case "48 Months Existing Business Acquisition & Retention Gas Tariff no S/C":
                    gasTariffId = "a0b1B00000P0lPX";
                    break;
                case "48 Months Existing Business Acquisition & Retention Gas Tariff":
                    gasTariffId = "a0b1B00000P0lPN";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdSE(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "renewalsmartfix – 1 rear level1":
                    gasTariffId = "a0b4v00000QDWa7";
                    break;
                case "renewalsmartfix – 2 year level1":
                    gasTariffId = "a0b4v00000QDWaM";
                    break;
                case "renewalsmartfix – 3 year level1":
                    gasTariffId = "a0b4v00000QDWaR";
                    break;
                case "renewalsmarttracker level1":
                    gasTariffId = "a0b4v00000QDWab";
                    break;
                case "acquisitionsmartfix – 1 year level1":
                    gasTariffId = "a0b4v00000QDWZs";
                    break;
                case "acquisitionsmartfix – 2 year level1":
                    gasTariffId = "a0b4v00000QDWZe";
                    break;
                case "acquisitionsmartfix – 3 year level1":
                    gasTariffId = "a0b4v00000QDWZx";
                    break;
                case "acquisitionsmarttracker level1":
                    gasTariffId = "a0b4v00000QDWaW ";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdSSE(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue)
            {
                case "12":
                    gasTariffId = "a0b1300000JQtcE";
                    break;
                case "24":
                    gasTariffId = "a0b1300000JQtcJ";
                    break;
                case "36":
                    gasTariffId = "a0b1300000JQtcT";
                    break;
                case "48":
                    gasTariffId = "a0b1300000JQtcY";
                    break;
                case "60":
                    gasTariffId = "a0b1B00000Q1UGL";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdCNG(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue.ToLower())
            {
                case "12no s/c":
                    gasTariffId = "a0b1B00000Q0qTi";
                    break;
                case "24no s/c":
                    gasTariffId = "a0b1B00000Q0qTs";
                    break;
                case "36no s/c":
                    gasTariffId = "a0b1B00000Q0qTx";
                    break;
                case "48no s/c":
                    gasTariffId = "a0b4v00000QDlX2";
                    break;
                case "12with s/c":
                    gasTariffId = "a0b30000002iZJH";
                    break;
                case "24with s/c":
                    gasTariffId = "a0b30000002jdoW";
                    break;
                case "36with s/c":
                    gasTariffId = "a0ba000000GdKld";
                    break;
                case "48with s/c":
                    gasTariffId = "a0b4v00000QDlX7";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdCG(string cellValue, string standingCharge)
        {
            string gasTariffId = string.Empty;

            if (standingCharge.Equals("0"))
            {
                switch (cellValue.ToLower())
                {
                    case "12":
                        gasTariffId = "a0b1300000MGLlR";
                        break;
                    case "24":
                        gasTariffId = "a0b1300000MGLlg";
                        break;
                    case "36":
                        gasTariffId = "a0b1300000MGLm5";
                        break;
                    case "48":
                        gasTariffId = "a0b1300000MGLmK";
                        break;
                    default:
                        gasTariffId = string.Empty;
                        break;
                }
            }
            else
            {
                switch (cellValue.ToLower())
                {
                    case "12":
                        gasTariffId = "a0b1300000LGZ2O";
                        break;
                    case "24":
                        gasTariffId = "a0b1300000LGZ2Y";
                        break;
                    case "36":
                        gasTariffId = "a0b1300000LGZ2d";
                        break;
                    case "48":
                        gasTariffId = "a0b1300000MGLmU";
                        break;
                    default:
                        gasTariffId = string.Empty;
                        break;
                }
            }

            return gasTariffId;
        }

        public string GetGasTariffIdDG_REN(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue.Substring(0, 2))
            {
                case "12":
                    gasTariffId = "a0b1B00000Q1LgH";
                    break;
                case "24":
                    gasTariffId = "a0b1B00000Q1LgM";
                    break;
                case "36":
                    gasTariffId = "a0b1B00000Q1LgR";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdDG_ACQ(string cellValue)
        {
            string gasTariffId = string.Empty;

            switch (cellValue.Substring(0, 2))
            {
                case "12":
                    gasTariffId = "a0b1B00000PSvLM";
                    break;
                case "24":
                    gasTariffId = "a0b1B00000PSvLR";
                    break;
                case "36":
                    gasTariffId = "a0b1B00000Q12FM";
                    break;
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public string GetGasTariffIdEON(string contactLength, string usageBand)
        {
            string gasTariffId = string.Empty;

            switch (usageBand)
            {
                case "0-24999":
                    {
                        if (contactLength.ToLower().Equals("12 months"))
                            gasTariffId = "a0b1300000Oufni";

                        break;
                    }
                case "0-14999":
                    {
                        if (contactLength.ToLower().Equals("12 months"))
                            gasTariffId = "a0b1B00000PYwfR";
                        else if (contactLength.ToLower().Equals("24 months"))
                            gasTariffId = "a0b1B00000PYwfb";
                        else if (contactLength.ToLower().Equals("36 months"))
                            gasTariffId = "a0b1B00000PYwfg";

                        break;
                    }
                case "15000-24999":
                    {
                        if (contactLength.ToLower().Equals("12 months"))
                            gasTariffId = "a0b1B00000PYwfv";
                        else if (contactLength.ToLower().Equals("24 months"))
                            gasTariffId = "a0b1B00000PYwg0";
                        else if (contactLength.ToLower().Equals("36 months"))
                            gasTariffId = "a0b1B00000PYwg5";
                        break;
                    }
                case "25000-54999":
                    {
                        if (contactLength.ToLower().Equals("12 months"))
                            gasTariffId = "a0b1300000Ouhfp";
                        else if (contactLength.ToLower().Equals("24 months"))
                            gasTariffId = "a0b1300000Ouhfu";
                        else if (contactLength.ToLower().Equals("36 months"))
                            gasTariffId = "a0b1300000Ouhfz";
                        break;
                    }
                case "55000-73267":
                    {
                        if (contactLength.ToLower().Equals("12 months"))
                            gasTariffId = "a0b1300000OuhgT";
                        else if (contactLength.ToLower().Equals("24 months"))
                            gasTariffId = "a0b1300000OuhgY";
                        else if (contactLength.ToLower().Equals("36 months"))
                            gasTariffId = "a0b1300000Ouhgd";
                        break;
                    }
                case "73268-99999999":
                    {
                        if (contactLength.ToLower().Equals("12 months"))
                            gasTariffId = "a0b1300000Oui0j";
                        else if (contactLength.ToLower().Equals("24 months"))
                            gasTariffId = "a0b1300000Oui0t";
                        else if (contactLength.ToLower().Equals("36 months"))
                            gasTariffId = "a0b1300000Oui0y";
                        break;
                    }
                default:
                    gasTariffId = string.Empty;
                    break;
            }

            return gasTariffId;
        }

        public int GetUsageBandMin(int cellValue)
        {
            int usageBandMin = 0;

            switch (cellValue)
            {
                case 4999:
                    usageBandMin = 0;
                    break;
                case 9999:
                    usageBandMin = 5000;
                    break;
                case 19999:
                    usageBandMin = 10000;
                    break;
                case 49999:
                    usageBandMin = 20000;
                    break;
                case 99999:
                    usageBandMin = 50000;
                    break;
                case 99999999:
                    usageBandMin = 100000;
                    break;
                default:
                    usageBandMin = 0;
                    break;
            }

            return usageBandMin;
        }

        public int GetUsageBandMinGas(int cellValue)
        {
            int usageBandMin = 0;

            switch (cellValue)
            {
                case 9999:
                    usageBandMin = 0;
                    break;
                case 19999:
                    usageBandMin = 10000;
                    break;
                case 39999:
                    usageBandMin = 20000;
                    break;
                case 73199:
                    usageBandMin = 40000;
                    break;
                default:
                    usageBandMin = 0;
                    break;
            }

            return usageBandMin;
        }

        public string GetProfileClassEDF(string cellValue)
        {
            string profileClass = string.Empty;

            switch (cellValue.ToLower())
            {
                case "std":
                    profileClass = "03";
                    break;
                case "ewe":
                    profileClass = "03";
                    break;
                case "ec7":
                    profileClass = "04";
                    break;
                case "ewn":
                    profileClass = "04";
                    break;
                default:
                    break;
            }

            return profileClass;
        }

        public string GetUniqueIdentifierBGL(Excel.Range range, int row)
        {
            string uniqueIdentifier = string.Empty;
            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 4]).Text;
            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 5]).Text;
            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 6]).Text;
            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 8]).Text;
            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 9]).Text;
            if (((Excel.Range)range.Cells[row, 10] != null) && (((Excel.Range)range.Cells[row, 10]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 10]).Text;
            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 11]).Text;
            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 2]).Text;
            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 3]).Text;

            return uniqueIdentifier;
        }

        public string GetUniqueIdentifierBG(Excel.Range range, int row)
        {
            string uniqueIdentifier = string.Empty;
            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 4]).Text;
            if (((Excel.Range)range.Cells[row, 5] != null) && (((Excel.Range)range.Cells[row, 5]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 5]).Text;
            if (((Excel.Range)range.Cells[row, 6] != null) && (((Excel.Range)range.Cells[row, 6]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 6]).Text;
            if (((Excel.Range)range.Cells[row, 8] != null) && (((Excel.Range)range.Cells[row, 8]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 8]).Text;
            if (((Excel.Range)range.Cells[row, 9] != null) && (((Excel.Range)range.Cells[row, 9]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 9]).Text;
            if (((Excel.Range)range.Cells[row, 11] != null) && (((Excel.Range)range.Cells[row, 11]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 11]).Text;
            if (((Excel.Range)range.Cells[row, 12] != null) && (((Excel.Range)range.Cells[row, 12]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 12]).Text;
            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 2]).Text;
            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 3]).Text;

            return uniqueIdentifier;
        }

        public string GetUniqueIdentifierOE(Excel.Range range, int row)
        {
            string uniqueIdentifier = string.Empty;
            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 4]).Text;

            return uniqueIdentifier;
        }

        public string GetUniqueIdentifierVE(Excel.Range range, int row)
        {
            string uniqueIdentifier = string.Empty;
            if (((Excel.Range)range.Cells[row, 1] != null) && (((Excel.Range)range.Cells[row, 1]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 1]).Text;
            if (((Excel.Range)range.Cells[row, 2] != null) && (((Excel.Range)range.Cells[row, 2]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 2]).Text;
            if (((Excel.Range)range.Cells[row, 3] != null) && (((Excel.Range)range.Cells[row, 3]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 3]).Text;
            if (((Excel.Range)range.Cells[row, 4] != null) && (((Excel.Range)range.Cells[row, 4]).Text != string.Empty))
                uniqueIdentifier += ((Excel.Range)range.Cells[row, 4]).Text;

            return uniqueIdentifier;
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