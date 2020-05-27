using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTracker.Tables
{
    class ProgTable
    {
        #region Properties
        public Guid ID { get; set; }
        public string WBS_ID { get; set; }
        public string ProgramTitle { get; set; }
        #endregion

        #region Constructor
        public ProgTable(string WBS_ID, string ProgramTitle)
        {
            this.ID = Guid.NewGuid();
            this.WBS_ID = WBS_ID;
            this.ProgramTitle = ProgramTitle;
        }
        #endregion
    }
}
