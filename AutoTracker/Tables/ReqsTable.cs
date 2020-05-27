using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTracker.Tables
{
    class ReqsTable
    {
        public Guid ID { get; set; }
        public Guid mapped_progID { get; set; }
        public string progName { get; set; }
        public string LRMK { get; set; }
        public string name { get; set; }
        public string totalReqs { get; set; }
        public string EN { get; set; }
        public string LG { get; set; }
        public string PK { get; set; }
        public string IN { get; set; }
        public string FM { get; set; }
        public string PM { get; set; }

        public ReqsTable(Guid mapped_progID, string progName, string LRMK, string name, string totalReqs, string EN, string LG, string PK, string IN, string FM, string PM)
        {
            this.ID = Guid.NewGuid();
            this.mapped_progID = mapped_progID;
            this.progName = progName;
            this.LRMK = LRMK;
            this.name = name;
            this.totalReqs = totalReqs;
            this.EN = EN;
            this.LG = LG;
            this.PK = PK;
            this.IN = IN;
            this.FM = FM;
            this.PM = PM;
        }
    }
}
