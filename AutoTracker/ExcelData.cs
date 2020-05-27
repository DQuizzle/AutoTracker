using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using LINQtoCSV;

namespace AutoTracker
{
    class ExcelData
    {
        [CsvColumn(Name = "WBS")]
        public string WBSCode { get; set; }

        [CsvColumn(Name = "GRADE")]
        public string GRADE { get; set; }

        [CsvColumn(Name = "OCC")]
        public string SERIES { get; set; }

        [CsvColumn(Name = "NAME")]
        public string NAME { get; set; }

        [CsvColumn(Name = "MPCN")]
        public string MPCN { get; set; }

        [CsvColumn(Name = "LRMK1")]
        public string LRMK { get; set; }

        [CsvColumn(Name = "WBS_TITLE")]
        public string WBS_TITLE { get; set; }


        public ExcelData() { }

        public ExcelData(string WBSCode, string GRADE, string SERIES, string NAME, string MPCN, string LRMK, string WBS_TITLE)
        {
            this.WBSCode = WBSCode;

            this.GRADE = GRADE;
            this.SERIES = SERIES;
            this.NAME = NAME;
            this.MPCN = MPCN;
            this.LRMK = LRMK;
            this.WBS_TITLE = WBS_TITLE;
        }
    }
}
