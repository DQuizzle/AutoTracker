using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Security.Cryptography;

namespace AutoTracker.Tables
{
    class ProgTable
    {
        #region Properties
        public Guid ID { get; set; }
        public string WBS_ID { get; set; }
        public string ProgramTitle { get; set; }
        public string Tier { get; set; }
        #endregion

        #region Constructor
        public ProgTable(string WBS_ID, string ProgramTitle, string Tier)
        {
            this.ID = Hash(ProgramTitle);
            this.WBS_ID = WBS_ID;
            this.ProgramTitle = ProgramTitle;
            this.Tier = Tier;
        }
        #endregion

        public Guid Hash(string progTitle)
        {
            using (MD5 md5 = MD5.Create())
            {
                byte[] inputBytes = Encoding.ASCII.GetBytes(progTitle);
                byte[] hashBytes = md5.ComputeHash(inputBytes);

                return new Guid(hashBytes);
            }
        }
    }
}
