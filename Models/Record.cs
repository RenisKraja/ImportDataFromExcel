using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ImportDataFromExcel.Models
{
    public class Record
    {
        public string GasTariffID
        {
            get;
            set;
        }

        public string LDZID
        {
            get;
            set;
        }

        public int UsageBandMin
        {
            get;
            set;
        }
    }
}