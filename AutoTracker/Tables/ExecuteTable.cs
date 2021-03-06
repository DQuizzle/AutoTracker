using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Security.Cryptography;

namespace AutoTracker.Tables
{
    class ExecuteTable
    {
        #region Properties
        public Guid mapped_ASUID { get; private set; }

        public string LRMK { get; private set; }
        public string Grade { get; private set; }
        public string Series { get; private set; }
        public string Name { get; private set; }
        public string MPCN { get; private set; }
        public string ProgramName { get; private set; }
        #endregion

        #region Constructor
        public ExecuteTable(Guid mapped_ASUID, string LRMK, string Grade, string Series, string Name, string MPCN, string programName)
        {
            this.mapped_ASUID = mapped_ASUID;
            this.LRMK = LRMK;
            this.Grade = Grade;
            this.Series = Series;
            this.Name = Name;
            this.MPCN = MPCN;
            this.ProgramName = programName;
        }
        #endregion
    }
}
