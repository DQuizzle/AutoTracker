using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTracker.Tables
{
    class UMDTable
    {
        #region Properties
        public Guid ID { get; private set; }

        public Guid mapped_ASUID { get; private set; }

        public string LRMK { get; private set; }
        public string Grade { get; private set; }
        public string Series { get; private set; }
        public string Name { get; private set; }
        public string MPCN { get; private set; }
        #endregion

        #region Constructor
        public UMDTable(Guid mapped_ASUID, string LRMK, string Grade, string Series, string Name, string MPCN)
        {
            this.ID = Guid.NewGuid();
            this.mapped_ASUID = mapped_ASUID;
            this.LRMK = LRMK;
            this.Grade = Grade;
            this.Series = Series;
            this.Name = Name;
            this.MPCN = MPCN;
        }
        #endregion
    }
}
