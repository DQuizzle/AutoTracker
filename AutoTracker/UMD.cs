using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using LINQtoCSV;

namespace AutoTracker
{
    class UMD
    {
        [CsvColumn(Name = "WBS")]
        public string WBS { get; set; }

        [CsvColumn(Name = "NAMES")]
        public string NAME { get; set; }

        [CsvColumn(Name = "AUTH_GRADE")]
        public string GRADE { get; set; }

        [CsvColumn(Name = "OCC")]
        public string SERIES { get; set; }

        [CsvColumn(Name = "MPCN")]
        public string MPCN { get; set; }

        [CsvColumn(Name = "LRMK1")]
        public string LRMK { get; set; }

        [CsvColumn(Name = "WBS_TITLE")]
        public string WBS_TITLE { get; set; }



        [CsvColumn(Name = "FUNDING_DESC")]
        public string FUNDING { get; set; }



        public UMD() { }

        public UMD(string WBS, string NAME, string GRADE, string SERIES, string MPCN, string LRMK, string FUNDING, string WBS_TITLE)
        {
            this.WBS = WBS;
            this.NAME = NAME;
            this.GRADE = GRADE;
            this.SERIES = SERIES;
            this.MPCN = MPCN;
            this.LRMK = LRMK;
            this.FUNDING = FUNDING;
            this.WBS_TITLE = WBS_TITLE;
        }
    }
}
