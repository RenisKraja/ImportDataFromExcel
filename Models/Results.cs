using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ImportDataFromExcel.Models
{
    public class Results
    {
        public string Status
        {
            get;
            set;
        }

        public string Object
        {
            get;
            set;
        }

        public string RecordCreated
        {
            get;
            set;
        }

        public string RecordFailed
        {
            get;
            set;
        }

        public string StartDate
        {
            get;
            set;
        }

        public string ProcessingTime
        {
            get;
            set;
        }

        public string MessageError
        {
            get;
            set;
        }

        public string StatusCode
        {
            get;
            set;
        }

        public string ReferenceId
        {
            get;
            set;
        }
    }
}