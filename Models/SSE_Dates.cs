using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ImportDataFromExcel.Models
{
    public class SSE_Dates
    {
        [DisplayName("First Date")]
        [DataType(DataType.Date)]
        public DateTime FirstDate
        {
            get;
            set;
        }

        [DisplayName("Second Date")]
        [DataType(DataType.Date)]
        public DateTime SecondDate
        {
            get;
            set;
        }

        [DisplayName("Third Date")]
        [DataType(DataType.Date)]
        public DateTime ThirdDate
        {
            get;
            set;
        }

        [DisplayName("Fourth Date")]
        [DataType(DataType.Date)]
        public DateTime FourthDate
        {
            get;
            set;
        }

        [DisplayName("Earliest Contract Start Date")]
        [DataType(DataType.Date)]
        public DateTime EarliestContractStartDate_First
        {
            get;
            set;
        }

        [DisplayName("Latest Contract Start Date")]
        [DataType(DataType.Date)]
        public DateTime LatestContractStartDate_First
        {
            get;
            set;
        }

        [DisplayName("Earliest Contract Start Date")]
        [DataType(DataType.Date)]
        public DateTime EarliestContractStartDate_Second
        {
            get;
            set;
        }

        [DisplayName("Latest Contract Start Date")]
        [DataType(DataType.Date)]
        public DateTime LatestContractStartDate_Second
        {
            get;
            set;
        }

        [DisplayName("Earliest Contract Start Date")]
        [DataType(DataType.Date)]
        public DateTime EarliestContractStartDate_Third
        {
            get;
            set;
        }

        [DisplayName("Latest Contract Start Date")]
        [DataType(DataType.Date)]
        public DateTime LatestContractStartDate_Third
        {
            get;
            set;
        }

        [DisplayName("Earliest Contract Start Date")]
        [DataType(DataType.Date)]
        public DateTime EarliestContractStartDate_Fourth
        {
            get;
            set;
        }

        [DisplayName("Latest Contract Start Date")]
        [DataType(DataType.Date)]
        public DateTime LatestContractStartDate_Fourth
        {
            get;
            set;
        }
    }
}