using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace AutoTracker
{
    public partial class Import : Form
    {
        #region Initialize Variables
        ExcelUtility util = new ExcelUtility();
        DataSet result = new DataSet();
        
        public Import()
        {
            InitializeComponent();
        }
        #endregion

        #region General Methods
        private void DeleteTempFiles()
        {
            File.Delete("temp.csv");
            File.Delete("temp1.csv");
            File.Delete("temp2.csv");
        }

        private void buttonChk()
        {
            button1.Enabled = (!String.IsNullOrEmpty(inputBox1.Text) && !String.IsNullOrEmpty(inputBox2.Text));
        }
        #endregion

        #region Button Methods and Operations
        private void btn_Browse1_Click(object sender, EventArgs e)
        {
            inputBox1.Text = util.Browse_xlsx();
            buttonChk();
        }

        private void btn_Browse2_Click(object sender, EventArgs e)
        {
            inputBox2.Text = util.Browse_xlsx();
            buttonChk();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            //Checks to see if Generating the CSVs were successful
            //If successful, opens a save dialog for user to select
            //where to save the .XML database file
            if (util.GenerateCSV(inputBox1.Text, inputBox2.Text, outputBox1.Text))
            {

                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "XML (*.xml)|*.xml";
                dialog.Title = "Select a location to save your database file.";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    newFileTextBox.Text = dialog.FileName;

                    //Executes the Parse to create the XML database file
                    ExcelParse.Reset();
                    if (ExcelParse.Parser("temp.csv"))
                    {
                        ExcelParse.WriteXML(newFileTextBox.Text);
                        ExcelParse.XMLPath = newFileTextBox.Text;
                        this.Cursor = Cursors.Default;
                        this.DialogResult = DialogResult.OK;

                        DeleteTempFiles();

                        this.Close();
                    }
                }
                else
                {
                    DeleteTempFiles();
                    return;
                }
            }

            this.Cursor = Cursors.Default;
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion
        
    }
}
