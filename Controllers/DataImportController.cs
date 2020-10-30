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

        public const string LoginEndpoint = "https://test.salesforce.com/services/oauth2/token"; //https://login.salesforce.com/services/oauth2/token
        public const string ApiEndpoint = "/services/data/v36.0/";//"/services/data/00D030000008aiM/";
        public string AuthToken = "";
        public string ServiceUrl = "";

        static HttpClient Client;

        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase exelFile)
        {
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

            //CreateAccount(Client);
            GetAccount();

            if ((exelFile == null) || (exelFile.ContentLength == 0))
            {
                ViewBag.Error = "Please select an excel file!";
                return View("Index");
            }
            else
            {
                if ((exelFile.FileName.EndsWith("xls")) || (exelFile.FileName.EndsWith("xlsx")) || (exelFile.FileName.EndsWith("csv")))
                {
                    string path = Server.MapPath("~/Content/" + exelFile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    exelFile.SaveAs(path);

                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workBook = application.Workbooks.Open(path);
                    Excel.Worksheet workSheet = workBook.ActiveSheet;
                    Excel.Range range = workSheet.UsedRange;
                    List<Record> listRecords = new List<Record>();
                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        CreateObjectType1(Client, range, row);
                        Record item = new Record();
                        //if (range.Cells[row, 11].Value2 != null)
                        if ((Excel.Range)range.Cells[row, 1] != null)
                            item.GasTariffID = ((Excel.Range)range.Cells[row, 1]).Text;
                        if ((Excel.Range)range.Cells[row, 2] != null)
                            item.LDZID = ((Excel.Range)range.Cells[row, 2]).Text;
                        if ((Excel.Range)range.Cells[row, 3] != null)
                            item.UsageBandMin = int.Parse(((Excel.Range)range.Cells[row, 3]).Text);

                        listRecords.Add(item);
                    }

                    ViewBag.ListRecords = listRecords;


                    workBook.Close(true, null, null);
                    application.Quit();

                    Marshal.ReleaseComObject(workSheet);
                    Marshal.ReleaseComObject(workBook);
                    Marshal.ReleaseComObject(application);

                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "File type is incorrect! <br>";
                    return View("Index");
                }
            }
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

        public void GetAccount()
        {
            string companyName = "Test123";
            string queryMessage = $"SELECT Id, Name, Phone, Type FROM Account WHERE Name = '{companyName}'";

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

    }
}