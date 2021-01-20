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

namespace ImportDataFromExcel.Models
{
    public class Methods
    {
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
                case "evening / weekend / night rate":
                    unitTypeFieldName = "Weekend_Rate__c";
                    break;
                case "evening / weekend rate":
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
    }
}