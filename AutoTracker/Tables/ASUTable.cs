using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTracker.Tables
{
    class ASUTable
    {
        #region Properties
        public Guid ID { get; set; }
        public string WBS { get; set; }
        #endregion

        #region Constructor
        public ASUTable(string WBS)
        {
            this.ID = Guid.NewGuid();
            this.WBS = WBS;
        }
        #endregion
    }
}
