using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoTracker
{
    public partial class AdminAccess : Form
    {
        static bool good = false;

        public AdminAccess()
        {
            InitializeComponent();
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Okbtn_Click(object sender, EventArgs e)
        {
            if (passwordBox.Text == "PEGAdmin2423")
            {
                good = true;
                Close();
            }
            else
                MessageBox.Show("Wrong password. Please try again", "Wrong Password", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static bool passwordGood()
        {
            if (good == true)
                return true;
            else
                return false;
        }
    }
}
