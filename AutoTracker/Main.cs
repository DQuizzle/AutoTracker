using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AutoTracker.Properties;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace AutoTracker
{
    public partial class Main : Form
    {
        #region Initialize Variables
        private DataSet ds;
        bool loaded, templateLoaded;
        const int MAX_UMD_PER_SLIDE = 20;
        string templateLocation;

        string saveFile;
        #endregion

        #region Initial/Generic Methods
        public Main()
        {
            InitializeComponent();

            FormsVisible(false);
            RefreshData();
        }

       
        private void RefreshData()
        {
            if (ExcelParse.MainDataSet == null)
                return;
            if (ExcelParse.MainDataSet.Tables["ASUTable"].Rows.Count != 0)
            {
                ds = ExcelParse.MainDataSet;
                aSUTableBindingSource.DataSource = ds;
                aSUTableBindingSource.Sort = "WBS Asc";

                aSUTableBindingSource1.DataSource = ds;
                aSUTableBindingSource1.Sort = "WBS Asc";

                uMDTableBindingSource.DataSource = ds;
                uMDTableBindingSource.Sort = "Name Asc";
                
                comboBox3.DataSource = new int[]{ 1, 2, 3, 4, 5 };

                comboBox1.SelectedIndex = 1;
                umdWBS.SelectedIndex = 1;
                dataGridView1.ClearSelection();
                
                FormsVisible(true);
                
                asuEditBox.Visible = true;
                deselectBtn.Enabled = true;
                exportBtn.Enabled = true;
                exportToolStripMenuItem.Enabled = true;
                umdGRB.Visible = false;
            }
        }

        private void FormsVisible(bool enable)
        {
            dataGridView1.Visible = enable;
            dataGridView2.Visible = enable;
            dataGridView3.Visible = enable;
            exec_Box.Visible = enable;
        }

        private void ClearReq()
        {
            name_txt.Text = "";
            totalReq_txt.Text = "";
            en_txt.Text = "";
            lg_txt.Text = "";
            pk_txt.Text = "";
            in_txt.Text = "";
            fm_txt.Text = "";
            pm_txt.Text = "";
        }

        private void ClearUMD()
        {
            gradeTxt.Text = "";
            umdName.Text = "";
            mpcn_Txt.Text = "";
        }

        private DataRow getCurrentUMDRow()
        {
            if (progTableUMDTableBindingSource.Current == null)
                return null;

            return (progTableUMDTableBindingSource.Current as DataRowView).Row;
        }

        private DataRow getCurrentASURow()
        {
            if (progTableReqsTableBindingSource.Current == null)
                return null;

            return (progTableReqsTableBindingSource.Current as DataRowView).Row;
        }

        private DataRow getCurrentExecRow()
        {
            if (progTableExecuteTableBindingSource.Current == null)
                return null;

            return (progTableExecuteTableBindingSource.Current as DataRowView).Row;
        }

        private void updateASUBtn_Click(object sender, EventArgs e)
        {
            UpdateASUForm();
        }

        private void updateUMDBtn_Click(object sender, EventArgs e)
        {
            UpdateUMDForm();
        }

        private void deselectBtn_Click(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
            dataGridView2.ClearSelection();
            dataGridView3.ClearSelection();
            ClearReq();
            ClearUMD();
            addASUBtn.Enabled = false;
            updateASUBtn.Enabled = false;
            delASUBtn.Enabled = false;
            addUMDBtn.Enabled = false;
            updateUMDBtn.Enabled = false;
            delUMD.Enabled = false;
            addExecuteBtn.Enabled = false;
            execDelBtn.Enabled = false;
        }

        private void dataGridView3_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView3.ClearSelection();
        }

        private void SetTitle(string path)
        {
            this.Text = "ASU Presenter - " + Path.GetFileNameWithoutExtension(path);
        }
        #endregion

        #region Cell Color Formatting Methods
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (Convert.ToString(row.Cells["nameDataGridViewTextBoxColumn1"].Value) == "Funded")
                {
                    row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#00CC00");
                }
                else if (Convert.ToString(row.Cells["nameDataGridViewTextBoxColumn1"].Value) == "Unfunded")
                {
                    row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#66FF33");
                }
                else if (Convert.ToString(row.Cells["nameDataGridViewTextBoxColumn1"].Value) == "CME")
                {
                    row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#FFC000");
                }
                else if (Convert.ToString(row.Cells["nameDataGridViewTextBoxColumn1"].Value) == "PEO Estimate")
                {
                    row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#C4D79B");
                }
                else if (Convert.ToString(row.Cells["nameDataGridViewTextBoxColumn1"].Value) == "Funded AMR" || Convert.ToString(row.Cells["nameDataGridViewTextBoxColumn1"].Value) == "AMR Funded")
                {
                    row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#AE78D6");
                }
                else if (Convert.ToString(row.Cells["nameDataGridViewTextBoxColumn1"].Value) == "" || Convert.ToString(row.Cells["nameDataGridViewTextBoxColumn1"].Value) == "Positions in Holding Cell (.1)")
                {
                    row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#BFBFBF");
                }
                else
                {
                    row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml("#B7DEE8");
                }
            }
        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row != null)
                {
                    if (Convert.ToString(row.Cells["gradeDataGridViewTextBoxColumn"].Value).Contains("GS") || Convert.ToString(row.Cells["gradeDataGridViewTextBoxColumn"].Value).Contains("NH")
                        || Convert.ToString(row.Cells["gradeDataGridViewTextBoxColumn"].Value).Contains("CME"))
                    {
                        row.DefaultCellStyle.BackColor = Color.LightGreen;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                }
            }
        }

        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                if (row != null)
                {
                    if (Convert.ToString(row.Cells["Grade"].Value).Contains("GS") || Convert.ToString(row.Cells["Grade"].Value).Contains("NH"))
                    {
                        row.DefaultCellStyle.BackColor = Color.LightGreen;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                }
            }
        }
        #endregion

        #region Tool Strip Menu Methods
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkSave("Are you sure you want to save ALL changes?");
        }

        private void importToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExcelParse.Reset();
            Import show = new Import();
            if (show.ShowDialog() == DialogResult.OK)
            {
                loaded = true;
                FormsVisible(true);
                saveFile = ExcelParse.XMLPath;

                SetTitle(saveFile);
                RefreshData();
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "XML (*.xml)|*.xml";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                checkSave("Would you like the save ALL changes?");

                DataSet newSet = new DataSet();
                try
                {
                    newSet.ReadXmlSchema(ExcelParse.XSDPath);
                    newSet.ReadXml(dialog.FileName);
                    ExcelParse.MainDataSet = newSet;
                    loaded = true;
                    FormsVisible(true);
                    ExcelParse.XMLPath = dialog.FileName;
                    saveFile = dialog.FileName;

                    SetTitle(saveFile);
                    RefreshData();
                }
                catch (Exception)
                {
                    MessageBox.Show("The XML file is corrupted, please re-import files", "Error opening file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkSave("Do you want to save ALL changes before closing program?");
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About openAbout = new About();
            openAbout.ShowDialog();
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            checkSave("Are you sure you want to save ALL changes?");
        }
        
        private void setTemplateLocationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Title = "Open Microsoft PowerPoint Template File";
            open.Filter = "POTX (*.potx)|*.potx";
            if (open.ShowDialog() == DialogResult.OK)
            {
                templateLocation = open.FileName;
                templateLoaded = true;
            }
            else
                templateLoaded = false;
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (saveBtn.Enabled == true)
            {
                DialogResult result = MessageBox.Show("Do you want to save ALL changes before closing program?", "Save Changes", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
                
                if (result == DialogResult.Yes)
                {
                    ExcelParse.WriteXML(saveFile);
                    
                    Environment.Exit(0);
                }
                else if (result == DialogResult.No)
                    Environment.Exit(0);
                else
                    e.Cancel = true;
            }
        }
        
        private void checkSave(string msg)
        {
            if (saveBtn.Enabled == true)
            {
                DialogResult result = MessageBox.Show(msg, "Save Changes", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

                if (result == DialogResult.Yes)
                {
                    ExcelParse.WriteXML(saveFile);

                    if (msg == "Do you want to save ALL changes before closing program?")
                        Environment.Exit(0);

                    saveBtn.Enabled = false;
                    saveToolStripMenuItem.Enabled = false;

                    SetTitle(saveFile);
                }
                else if (result == DialogResult.No)
                {
                    if (msg == "Do you want to save ALL changes before closing program?")
                        Environment.Exit(0);
                }
                else
                    return;
            }
        }
        #endregion

        #region Form Modifying Methods
        private void UpdateASUForm()
        {
            int i = dataGridView1.CurrentRow.Index;

            if (name_txt.Text != "" && totalReq_txt.Text != "" && en_txt.Text != "" && lg_txt.Text != ""
                && pk_txt.Text != "" && in_txt.Text != "" && fm_txt.Text != "" && pm_txt.Text != "")
            {
                dataGridView1.Rows[i].Cells[0].Value = name_txt.Text;
                dataGridView1.Rows[i].Cells[1].Value = totalReq_txt.Text;
                dataGridView1.Rows[i].Cells[2].Value = en_txt.Text;
                dataGridView1.Rows[i].Cells[3].Value = lg_txt.Text;
                dataGridView1.Rows[i].Cells[4].Value = pk_txt.Text;
                dataGridView1.Rows[i].Cells[5].Value = in_txt.Text;
                dataGridView1.Rows[i].Cells[6].Value = fm_txt.Text;
                dataGridView1.Rows[i].Cells[7].Value = pm_txt.Text;
                saveBtn.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                updateASUBtn.Enabled = false;
                addASUBtn.Enabled = false;
            }
            else
            {
                MessageBox.Show("Please ensure no entries are empty.", "Error Modifying Values", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateUMDForm()
        {
            int i = dataGridView2.CurrentRow.Index;

            if (gradeTxt.Text != "" && umdName.Text != "" && mpcn_Txt.Text != "" )
            {
                dataGridView2.Rows[i].Cells[0].Value = gradeTxt.Text;
                dataGridView2.Rows[i].Cells[1].Value = umdName.Text;
                dataGridView2.Rows[i].Cells[2].Value = mpcn_Txt.Text;
                saveBtn.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                updateUMDBtn.Enabled = false;
                addUMDBtn.Enabled = false;
            }
            else
            {
                MessageBox.Show("Please ensure no entries are empty.", "Error Modifying Values", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RemoveUMDRow(List<DataRow> drSelected)
        {
            string ID, name, MPCN;
            DataRow[] drUMDList;

            DataTable dtUMD = ExcelParse.MainDataSet.Tables["UMDTable"];

            foreach (DataRow dr in drSelected)
            {
                ID = dr["ID"].ToString();
                name = dr["Name"].ToString();
                MPCN = dr["MPCN"].ToString();
                drUMDList = dtUMD.Select("MPCN='" + MPCN + "'");

                DialogResult result = MessageBox.Show("Are you sure you want to delete " + name + " from the database?", "Delete from Database", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    DeleteUMDRow(dr);
                    saveBtn.Enabled = true;
                    saveToolStripMenuItem.Enabled = true;
                    addUMDBtn.Enabled = false;
                }
                else
                    break;
            }
        }

        public void DeleteUMDRow(DataRow drUMD)
        {
            DataTable dtUMD = ExcelParse.MainDataSet.Tables["UMDTable"];
            DataRow[] drUMDList = dtUMD.Select("MPCN='" + drUMD["MPCN"] + "'");

            foreach (DataRow drUMDPerson in drUMDList)
            {
                drUMDPerson.Delete();
            }

            drUMD.Delete();
        }

        private void RemoveReqRow(List<DataRow> drSelected)
        {
            string ID, name;
            DataRow[] drReqList;

            DataTable dtReqList = ExcelParse.MainDataSet.Tables["ReqsTable"];

            foreach (DataRow dr in drSelected)
            {
                ID = dr["ID"].ToString();
                name = dr["Name"].ToString();
                drReqList = dtReqList.Select("ID='" + ID + "'");

                DialogResult result = MessageBox.Show("Are you sure you want to delete " + name + " from the database?", "Delete from Database", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    DeleteReqRow(dr);
                    saveBtn.Enabled = true;
                    saveToolStripMenuItem.Enabled = true;
                    addUMDBtn.Enabled = false;
                }
                else
                    break;
            }
        }

        public void DeleteReqRow(DataRow drReq)
        {
            DataTable dtReq = ExcelParse.MainDataSet.Tables["ReqsTable"];
            DataRow[] drReqList = dtReq.Select("ID='" + drReq["ID"] + "'");

            foreach (DataRow row in drReqList)
            {
                row.Delete();
            }

            drReq.Delete();
        }

        private void RemoveExecRow(List<DataRow> drSelected)
        {
            string name, MPCN;
            DataRow[] drExecList;

            DataTable dtExec = ExcelParse.MainDataSet.Tables["ExecuteTable"];

            foreach (DataRow dr in drSelected)
            {
                name = dr["Name"].ToString();
                MPCN = dr["MPCN"].ToString();
                drExecList = dtExec.Select("MPCN='" + MPCN + "'");

                DialogResult result = MessageBox.Show("Are you sure you want to delete " + name + " from the Execute Table?", "Delete from Database", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    DeleteExecRow(dr);
                    saveBtn.Enabled = true;
                    saveToolStripMenuItem.Enabled = true;
                    execDelBtn.Enabled = false;
                }
                else
                    break;
            }
        }

        public void DeleteExecRow(DataRow drExec)
        {
            DataTable dtExec = ExcelParse.MainDataSet.Tables["ExecuteTable"];
            DataRow[] drExecList = dtExec.Select("MPCN='" + drExec["MPCN"] + "' AND PROG_ID='" + drExec["PROG_ID"] + "'");

            DataTable dtUMDExec = ExcelParse.MainDataSet.Tables["UMDTable"];
            DataRow[] drUMDExecList = dtUMDExec.Select("MPCN='" + drExec["MPCN"] + "' AND PROG_ID='" + drExec["PROG_ID"] + "'");

            foreach (DataRow drExecPerson in drExecList)
            {
                drExecPerson.Delete();
            }

            drExec.Delete();
        }

        private void addUMDBtn_Click(object sender, EventArgs e)
        {
            string reqProgID;

            if (dataGridView2.CurrentRow != null)
            {
                int i = dataGridView2.CurrentRow.Index;
                reqProgID = dataGridView2.Rows[i].Cells[3].Value.ToString();
            }
            else
                reqProgID = null;

            if (gradeTxt.Text != "" && umdName.Text != "" && mpcn_Txt.Text != "")
            {
                DataTable dt = ExcelParse.MainDataSet.Tables["UMDTable"];
                DataRow[] drMPCNChk = dt.Select("MPCN='" + mpcn_Txt.Text + "'");

                if (noMPCNChk.Checked == false)
                {
                    if (drMPCNChk.Length > 0)
                    {
                        MessageBox.Show("MPCN already exists.", "Existing MPCN", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        if (reqProgID == null)
                        {
                            int i = dataGridView1.CurrentRow.Index;
                            reqProgID = dataGridView1.Rows[i].Cells[8].Value.ToString();
                        }
                        DataRow newRow = dt.NewRow();

                        newRow["ID"] = Guid.NewGuid();
                        newRow["PROG_ID"] = reqProgID;
                        newRow["LRMK_ID"] = comboBox2.Text;
                        newRow["Grade"] = gradeTxt.Text;
                        newRow["Series"] = "";
                        newRow["Name"] = umdName.Text;
                        newRow["MPCN"] = mpcn_Txt.Text;

                        dt.Rows.Add(newRow);

                        saveBtn.Enabled = true;
                        saveToolStripMenuItem.Enabled = true;
                        addUMDBtn.Enabled = false;
                        updateUMDBtn.Enabled = false;
                   }
                }
                else
                {
                    if (reqProgID == null)
                    {
                        int i = dataGridView1.CurrentRow.Index;
                        reqProgID = dataGridView1.Rows[i].Cells[8].Value.ToString();
                    }
                    DataRow newRow = dt.NewRow();

                    newRow["ID"] = Guid.NewGuid();
                    newRow["PROG_ID"] = reqProgID;
                    newRow["LRMK_ID"] = comboBox2.Text;
                    newRow["Grade"] = gradeTxt.Text;
                    newRow["Series"] = "";
                    newRow["Name"] = umdName.Text;
                    newRow["MPCN"] = mpcn_Txt.Text;

                    dt.Rows.Add(newRow);

                    saveBtn.Enabled = true;
                    saveToolStripMenuItem.Enabled = true;
                    addUMDBtn.Enabled = false;
                    updateUMDBtn.Enabled = false;
                }
            }
            else
            {
                MessageBox.Show("Please ensure no entries are empty.", "Error Modifying Values", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void delASUBtn_Click(object sender, EventArgs e)
        {
            DataRow drASU = getCurrentASURow();
            if (drASU == null)
                return;

            List<DataRow> selectedRows = new List<DataRow>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Selected == true)
                    selectedRows.Add(((DataRowView)row.DataBoundItem).Row);
            }

            RemoveReqRow(selectedRows);
            dataGridView1.Refresh();
        }

        private void delUMD_Click(object sender, EventArgs e)
        {
            DataRow drUMD = getCurrentUMDRow();
            if (drUMD == null)
                return;

            List<DataRow> selectedRows = new List<DataRow>();

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.Selected == true)
                    selectedRows.Add(((DataRowView)row.DataBoundItem).Row);
            }

            RemoveUMDRow(selectedRows);
            dataGridView2.Refresh();
        }

        private void execDelBtn_Click(object sender, EventArgs e)
        {
            DataRow drExec = getCurrentExecRow();
            if (drExec == null)
                return;

            List<DataRow> selectedRows = new List<DataRow>();

            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                if (row.Selected == true)
                    selectedRows.Add(((DataRowView)row.DataBoundItem).Row);
            }

            RemoveExecRow(selectedRows);
            dataGridView3.Refresh();
        }

        private void addASUBtn_Click(object sender, EventArgs e)
        {
            string reqProgID;
            int i = 0;

            if (dataGridView1.CurrentRow != null)
            {
                i = dataGridView1.CurrentRow.Index;
                reqProgID = dataGridView1.Rows[i].Cells[8].Value.ToString();
            }
            else
                reqProgID = null;

            if (name_txt.Text != "" && totalReq_txt.Text != "" && en_txt.Text != "" && lg_txt.Text != ""
                && pk_txt.Text != "" && in_txt.Text != "" && fm_txt.Text != "" && pm_txt.Text != "")
            {
                DataTable dt = ExcelParse.MainDataSet.Tables["ReqsTable"];

                if (reqProgID != null)
                {
                    DataRow[] drProgIDs = dt.Select("ProgID='" + dataGridView1.Rows[i].Cells[8].Value.ToString() + "'");

                    foreach (DataRow row in drProgIDs)
                    {
                        string test = row["Name"].ToString();

                        if (test == name_txt.Text)
                        {
                            MessageBox.Show(name_txt.Text + " already exists.", "Existing " + name_txt.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }

                DataRow newRow = dt.NewRow();

                string name = name_txt.Text;
                string totalReq = totalReq_txt.Text;
                string en = en_txt.Text;
                string lg = lg_txt.Text;
                string pk = pk_txt.Text;
                string instring = in_txt.Text;
                string fm = fm_txt.Text;
                string pm = pm_txt.Text;

                if (reqProgID == null)
                {
                    reqProgID = dataGridView2.Rows[i].Cells[3].Value.ToString();
                }

                newRow["ID"] = Guid.NewGuid();
                newRow["ProgID"] = reqProgID;
                newRow["ProgName"] = comboBox2.Text;
                newRow["LRMK"] = comboBox2.Text;
                newRow["Name"] = name_txt.Text;
                newRow["TotalReqs"] = totalReq_txt.Text;
                newRow["EN"] = en_txt.Text;
                newRow["LG"] = lg_txt.Text;
                newRow["PK"] = pk_txt.Text;
                newRow["IN"] = in_txt.Text;
                newRow["FM"] = fm_txt.Text;
                newRow["PM"] = pm_txt.Text;

                dt.Rows.Add(newRow);
                
                saveBtn.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                addASUBtn.Enabled = false;
                updateASUBtn.Enabled = false;
            }
            else
            {
                MessageBox.Show("Please ensure no entries are empty.", "Error Modifying Values", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            dataGridView1.Refresh();
        }

        private void addExecTable(object reqProgID, DataTable newExecDt, DataRow[] drMPCNChk, string programName)
        {
            DataRow newRow = newExecDt.NewRow();

            newRow["PROG_ID"] = reqProgID;
            newRow["LRMK_ID"] = drMPCNChk[0].ItemArray[2];
            newRow["Grade"] = drMPCNChk[0].ItemArray[3];
            newRow["Series"] = drMPCNChk[0].ItemArray[4];
            newRow["Name"] = drMPCNChk[0].ItemArray[5];
            newRow["MPCN"] = drMPCNChk[0].ItemArray[6];
            newRow["ProgramName"] = programName;

            newExecDt.Rows.Add(newRow);

            saveBtn.Enabled = true;
            saveToolStripMenuItem.Enabled = true;
            updateUMDBtn.Enabled = false;
        }

        private void exec_AddBtn_Click(object sender, EventArgs e)
        {
            var reqProgID = comboBox2.SelectedValue;
            string programName;

            DataTable dt = ExcelParse.MainDataSet.Tables["UMDTable"];
            DataRow[] drMPCNChk = dt.Select("MPCN='" + execMPCN.Text + "'");
            
            DataTable findProgName = ExcelParse.MainDataSet.Tables["ProgTable"];
            DataRow[] drProgName = findProgName.Select("ID ='" + drMPCNChk[0].ItemArray[1].ToString() + "'");

            DataTable newExecDt = ExcelParse.MainDataSet.Tables["ExecuteTable"];
            DataRow[] drExecIDChk = newExecDt.Select("MPCN='" + execMPCN.Text + "'");
            
            if (drProgName == null)
                programName = comboBox2.Text;
            else
                programName = drProgName[0].ItemArray[2].ToString();

            if (drExecIDChk.Length > 0)
            {
                drExecIDChk = newExecDt.Select("MPCN='" + execMPCN.Text + "' AND PROG_ID='" + reqProgID + "'");

                if (drExecIDChk.Length > 0)
                {
                    MessageBox.Show(exec_NameBox.Text + " already assigned to this program.", "Existing Person", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                    addExecTable(reqProgID, newExecDt, drMPCNChk, programName);
            }
            else
            {
                addExecTable(reqProgID, newExecDt, drMPCNChk, programName);
            }
        }

        private void addExecuteBtn_Click(object sender, EventArgs e)
        {
            string progID, progName;

            DataTable selectProgram = ExcelParse.MainDataSet.Tables["ProgTable"];
            DataRow[] drSelectProgram = selectProgram.Select("ID='" + umdProgs.SelectedValue.ToString() + "'");

            DataTable dt = ExcelParse.MainDataSet.Tables["UMDTable"];
            DataRow[] drMPCNChk = dt.Select("MPCN='" + mpcn_Txt.Text + "'");
            
            DataRow[] drProgName = selectProgram.Select("ID='" + drMPCNChk[0].ItemArray[1].ToString() + "'");

            if (useCurrentProg.Checked == true)
            {
                progID = comboBox2.SelectedValue.ToString();
                progName = comboBox2.Text;
            }
            else
            {
                progID = drSelectProgram[0].ItemArray[0].ToString();
                progName = drProgName[0].ItemArray[2].ToString();
            }

            DataTable newExecDt = ExcelParse.MainDataSet.Tables["ExecuteTable"];
            DataRow[] drExecIDChk = newExecDt.Select("MPCN='" + mpcn_Txt.Text + "' AND PROG_ID='" + progID + "'");

            if (drExecIDChk.Length > 0)
            {
                MessageBox.Show(umdName.Text + " already assigned to this program.", "Existing Person", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                addExecTable(progID, newExecDt, drMPCNChk, progName);
            }
        }
        #endregion

        #region DataGrid Operation Methods
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            asuEditBox.Visible = true;
            umdGRB.Visible = false;

            if (e.RowIndex >= 0)
            {
                name_txt.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                totalReq_txt.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                en_txt.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                lg_txt.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                pk_txt.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                in_txt.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
                fm_txt.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
                pm_txt.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();

                addASUBtn.Enabled = false;
                updateASUBtn.Enabled = false;
                delASUBtn.Enabled = true;
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            asuEditBox.Visible = false;
            umdGRB.Visible = true;

            if (e.RowIndex >= 0)
            {
                gradeTxt.Text = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
                umdName.Text = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
                mpcn_Txt.Text = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();

                addUMDBtn.Enabled = false;
                updateUMDBtn.Enabled = false;
                delUMD.Enabled = true;
                addExecuteBtn.Enabled = true;
            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            execDelBtn.Enabled = true;
        }
        #endregion

        #region TextBox Validation Methods
        private void useCurrentProg_CheckedChanged(object sender, EventArgs e)
        {
            if (useCurrentProg.Checked == true)
            {
                umdWBS.Enabled = false;
                umdProgs.Enabled = false;
            }
            else
            {
                umdWBS.Enabled = true;
                umdProgs.Enabled = true;
            }
        }
        private void gradeTxt_KeyDown(object sender, KeyEventArgs e)
        {
            updateUMDBtn.Enabled = true;
            addUMDBtn.Enabled = true;
        }


        private void umdName_KeyDown(object sender, KeyEventArgs e)
        {
            updateUMDBtn.Enabled = true;
            addUMDBtn.Enabled = true;
        }

        private void mpcn_Txt_KeyDown(object sender, KeyEventArgs e)
        {
            updateUMDBtn.Enabled = true;
            addUMDBtn.Enabled = true;
        }

        private void name_txt_KeyDown(object sender, KeyEventArgs e)
        {
            updateASUBtn.Enabled = true;
            addASUBtn.Enabled = true;
        }

        private void totalReq_txt_KeyDown(object sender, KeyEventArgs e)
        {
            updateASUBtn.Enabled = true;
            addASUBtn.Enabled = true;
        }

        private void en_txt_KeyDown(object sender, KeyEventArgs e)
        {
            updateASUBtn.Enabled = true;
            addASUBtn.Enabled = true;
        }

        private void lg_txt_KeyDown(object sender, KeyEventArgs e)
        {
            updateASUBtn.Enabled = true;
            addASUBtn.Enabled = true;
        }

        private void pk_txt_KeyDown(object sender, KeyEventArgs e)
        {
            updateASUBtn.Enabled = true;
            addASUBtn.Enabled = true;
        }

        private void in_txt_KeyDown(object sender, KeyEventArgs e)
        {
            updateASUBtn.Enabled = true;
            addASUBtn.Enabled = true;
        }

        private void fm_txt_KeyDown(object sender, KeyEventArgs e)
        {
            updateASUBtn.Enabled = true;
            addASUBtn.Enabled = true;
        }

        private void pm_txt_KeyDown(object sender, KeyEventArgs e)
        {
            updateASUBtn.Enabled = true;
            addASUBtn.Enabled = true;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (asuEditBox.Visible == true)
                umdGRB.Visible = false;
            else
                umdGRB.Visible = true;

            ClearReq();
            ClearUMD();
            addExecuteBtn.Enabled = false;
            execDelBtn.Enabled = false;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (asuEditBox.Visible == true)
                umdGRB.Visible = false;
            else
                umdGRB.Visible = true;

            ClearReq();
            ClearUMD();
            addExecuteBtn.Enabled = false;
            execDelBtn.Enabled = false;
        }

        private void addNewUMD_Click(object sender, EventArgs e)
        {
            asuEditBox.Visible = false;
            umdGRB.Visible = true;
        }
        #endregion

        #region Generate PowerPoint
        private void PowerPoint(string path)
        {
            int cur_slide = 1;
            
            //Open template file, if not present, return to application
            if (!templateLoaded)
            {
                OpenFileDialog open = new OpenFileDialog();
                open.Title = "Open Microsoft PowerPoint Template File";
                open.Filter = "POTX (*.potx)|*.potx";
                if (open.ShowDialog() == DialogResult.OK)
                {
                    templateLocation = open.FileName;
                    templateLoaded = true;
                }
                else return;
            }
            
            PowerPoint.Application pptApplication = new PowerPoint.Application();

            PowerPoint.Slides slides;
            
            //Create the Presentation File
            PowerPoint.Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            PowerPoint.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];
            
            //Save PowerPoint
            try
            {
                pptPresentation.SaveAs(@path, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            }
            catch (Exception)
            {
                MessageBox.Show("ERROR: File is currently active in PowerPoint. Please close the PowerPoint file to export.", "Existing File Open", MessageBoxButtons.OK, MessageBoxIcon.Error);
                pptPresentation.Close();
                pptApplication.Quit();
                
                return;
            }
            
            //Create new Slide
            slides = pptPresentation.Slides;
            slides.AddSlide(1, customLayout);
            slides.AddSlide(2, customLayout);
            slides[cur_slide].ApplyTemplate(@templateLocation);
          
            //Create DataTables for the DataGridViews
            DataTable dt = CreateDataTableFromDGV(dataGridView1);
            DataTable dt_UMD = CreateDataTableFromDGV(dataGridView2);
            DataTable dt_Exec = CreateDataTableFromDGV(dataGridView3);
            
            //Generate Slides
            CreateSlide(slides, ref cur_slide, dt, dt_UMD, dt_Exec, true, customLayout);
            slides.AddSlide(cur_slide, customLayout);
            
            CreateSlide(slides, ref cur_slide, dt, dt_UMD, dt_Exec, false, customLayout);
            
            slides[cur_slide].Delete();
        }
        
        private void CreateSlide(PowerPoint.Slides slides, ref int cur_slide, DataTable dt, DataTable dt_UMD, DataTable dt_Exec, bool isUMD, PowerPoint.CustomLayout customLayout)
        {
            int itr = 0;
            DataTable dtParser;
            
            if (isUMD)
                dtParser = dt_UMD;
            else
                dtParser = dt_Exec;
                
            if (dtParser.Rows.Count > MAX_UMD_PER_SLIDE)
            {
                while (itr < dtParser.Rows.Count)
                {
                    SetupUMDorExecute(isUMD, slides, cur_slide, dt, dt_UMD, dt_Exec, ref itr);
                    
                    if (itr < dtParser.Rows.Count)
                    {
                        cur_slide++;
                        slides.AddSlide(cur_slide, customLayout);
                        slides[cur_slide].ApplyTemplate(@templateLocation);
                    }
                }
            }
            else
                SetupUMDorExecute(isUMD, slides, cur_slide, dt, dt_UMD, dt_Exec, ref itr);
                
            cur_slide++;
        }
        
        private void SetupUMDorExecute(bool isUMD, PowerPoint.Slides slides, int slide_num, DataTable dt, DataTable dt_UMD, DataTable dt_Exec, ref int itr)
        {
            if (isUMD)
            {
                //Add Title
                AddPPHeader(slides[slide_num], "");
            
                //Create ASU Table
                CreateASUTableLayout(slides[slide_num], dt);
            
                //Create UMD Table
                if (dataGridView2.RowCount != 0)
                    SetUpLowerPPTable(slides[slide_num], ref itr, dt_UMD);
                
                //Setup Tier Data
                SetUpTierDataTable(slides[slide_num], dt, dt_UMD, dt_Exec, 1, 680);
            
                //Setup LEGEND
                SetUpPPLegend(slides[slide_num], 120, 2);
            
                //Notes on PowerPoint Slides
                slides[slide_num].NotesPage.Shapes[2].TextFrame.TextRange.Text = "Funded Table";
            }
            else
            {
                //Add Title and customize Execution Table
                slides[slide_num].ApplyTemplate(@templateLocation);
                AddPPHeader(slides[slide_num], " (EXECUTION)");
                
                //Create Execution Table
                if (dataGridView3.RowCount != 0)
                    SetUpLowerExecPPTable(slides[slide_num], ref itr, dt_Exec);
                    
                //Setup Tier Data
                SetUpTierDataTable(slides[slide_num], dt, dt_UMD, dt_Exec, 2, 720);
            
                //Setup LEGEND
                SetUpPPLegend(slides[slide_num], 450, 3);
                
                //Delete blank placeholder behind Execution Data
                slides[slide_num].Shapes.Placeholders[2].Delete();
            
                //Notes on PowerPoint Slides
                slides[slide_num].NotesPage.Shapes[2].TextFrame.TextRange.Text = "Executed Table";
            }
        }

        private void AddPPHeader(PowerPoint._Slide slide, string add)
        {
            PowerPoint.TextRange objText;

            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = comboBox2.Text + "\n" + comboBox1.Text + add;
            objText.Font.Name = "Arial";
            objText.Font.Color.RGB = ColorTranslator.ToOle(Color.DarkBlue);
            objText.Font.Bold = MsoTriState.msoTrue;
            if (comboBox2.Text.Length > 34)
                objText.Font.Size = 22;
            else
                objText.Font.Size = 28;

            var objUnclass = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 680, 0, 100, 20);
            objUnclass.TextFrame.TextRange.Font.Size = 8;
            objUnclass.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            objUnclass.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Red);
            objUnclass.TextFrame.TextRange.Text = "UNCLASSIFIED";
        }

        private void CreateASUTableLayout(PowerPoint._Slide slide, DataTable dt)
        {
            if (dataGridView1.RowCount != 0)
            {
                var objShape = slide.Shapes.AddTable(dt.Rows.Count + 1, dt.Columns.Count - 2);
                objShape.Table.Columns._Index(1).Width = 280;
                objShape.Table.Columns._Index(2).Width = 64;
                for (int i = 3; i <= objShape.Table.Columns.Count; i++)
                {
                    objShape.Table.Columns._Index(i).Width = 40;
                }
                for (int i = 1; i <= objShape.Table.Rows.Count; i++)
                {
                    objShape.Table.Rows._Index(i).Height = 20;
                }
                var table = objShape.Table;

                SetUpASUPPTable(table, dt);
            }
            else
            {
                Microsoft.Office.Interop.PowerPoint.TextRange objNoASU;

                objNoASU = slide.Shapes[2].TextFrame.TextRange;
                objNoASU.Text = "NO ASU DATA AVAILABLE";
            }
        }

        private void SetUpPPLegend(PowerPoint._Slide slide, int y_pass, int count)
        {
            string[] name = { "= CIV / CME", "= MILITARY", "= NOT IN PROGRAM" };
            int x = 800;
            int y = y_pass;
            int l = 15;
            int h = 15;

            var objTitle = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, x - 4, y - 27, 100, h);
            objTitle.TextFrame.TextRange.Font.Size = 13;
            objTitle.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            objTitle.TextFrame.TextRange.Text = "LEGEND";

            for (int i = 0; i < count; i++)
            {
                var objRectangle = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, l, h);
                var objText = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, x + 12, y - 2, 100, h);

                objRectangle.Line.Weight = 1;

                objText.TextFrame.TextRange.Font.Size = 11;
                objText.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;

                if (i == 0)
                {
                    objRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGreen);
                    objText.TextFrame.TextRange.Text = name[i];
                }
                else if (i == 1)
                {
                    objRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
                }
                else
                {
                    objRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGoldenrodYellow);
                    objRectangle.Line.DashStyle = MsoLineDashStyle.msoLineDash;
                    objText.TextFrame.TextRange.Text = name[i];
                }

                y += 20;
            }
        }

        private void SetUpASUPPTable(PowerPoint.Table table, DataTable dt)
        {
            table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "";
            table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "TotalReqs";
            table.Cell(1, 3).Shape.TextFrame.TextRange.Text = "EN";
            table.Cell(1, 4).Shape.TextFrame.TextRange.Text = "LG";
            table.Cell(1, 5).Shape.TextFrame.TextRange.Text = "PK";
            table.Cell(1, 6).Shape.TextFrame.TextRange.Text = "IN";
            table.Cell(1, 7).Shape.TextFrame.TextRange.Text = "FM";    
            table.Cell(1, 8).Shape.TextFrame.TextRange.Text = "PM";

            for (int i = 2; i < 9; i++)
            {
                table.Cell(1, i).Shape.TextFrame.TextRange.Font.Size = 12;
                table.Cell(1, i).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                table.Cell(1, i).Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
            }

            table.Cell(1, 2).Shape.TextFrame.TextRange.Font.Size = 10;

            for (int i = 2; i <= table.Rows.Count; i++)
            {
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    table.Cell(i, j).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    table.Cell(i, j).Shape.TextFrame.TextRange.Font.Size = 12;
                    table.Cell(i, j).Shape.TextFrame.TextRange.Text = dt.Rows[i - 2].ItemArray[j - 1].ToString(); ;
                }
            }
        }
        
        private float getLatestFYData(DataTable dt)
        {
            var year = int.Parse(DateTime.Now.ToString("yy"));
            
            while (year != 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i].ItemArray[0].ToString().Contains(year.ToString()))
                        return float.Parse(dt.Rows[i].ItemArray[1].ToString());
                        
                    else if (dt.Rows[i].ItemArray[0].ToString().Contains("PEO"))
                        return float.Parse(dt.Rows[i].ItemArray[1].ToString());
                }
                
                year -= 1;
            }
            
            return float.Parse(dt.Rows[0].ItemArray[1].ToString());
        }
        
        private void SetUpTierDataTable(PowerPoint._Slide slide, DataTable dtASU, DataTable dtUMD, DataTable dtExec, int count, int x_pass)
        {
            int x = x_pass, y = 100, w = 90, h = 50, standard = 0;
            
            float filledCount = 0, actualCount = 0;
            int[] filled = { 0, 0 };
            string[] title = { "DATA", "ACTUAL" };
            
            float totalReqs = getLatestFYData(dtASU);
            int funded = (int)Math.Round((dtUMD.Rows.Count / totalReqs) * 100);
            
            for (int i = 0; i < dtUMD.Rows.Count; i++)
            {
                if (dtUMD.Rows[i].ItemArray[1].ToString() != "VACANT")
                    filledCount++;
            }
            
            filled[0] = (int)Math.Round((filledCount / totalReqs) * 100);
            
            for (int i = 0; i < dtExec.Rows.Count; i++)
            {
                if (dtExec.Rows[i].ItemArray[1].ToString() != "VACANT")
                    actualCount++;
            }
            
            filled[1] = (int)Math.Round((actualCount / totalReqs) * 100);
            
            switch (int.Parse(comboBox3.Text))
            {
                case 1:
                    standard = 90;
                    break;
                case 2:
                    standard = 88;
                    break;
                case 3:
                    standard = 85;
                    break;
                case 4:
                    standard = 82;
                    break;
                case 5:
                    standard = 80;
                    break;
            }
            
            for (int i = 0; i < count; i++)
            {
                var objRectangle = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, w, h);
                objRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.PowderBlue);
                objRectangle.TextFrame.TextRange.Font.Size = 11;
                objRectangle.TextFrame.TextRange.Text = title[i] + "\nTIER: " + comboBox3.Text + "\nStandard: " + standard + "%\nFunded: "
                    + funded + "%\nFilled: " + filled[i] + "%";
                objRectangle.TextFrame.TextRange.Paragraphs(1).Lines(1, 2).Font.Bold = MsoTriState.msoTrue;
                objRectangle.TextFrame.TextRange.Paragraphs(1).Lines(1, 2).Font.Size = 13;
                objRectangle.TextFrame.TextRange.Paragraphs(1).Lines(1).Font.Underline = MsoTriState.msoTrue;
                objRectangle.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                
                if (i == 1)
                    objRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Pink);
                    
                x += 100;
            }
        }

        private void SetUpLowerPPTable(PowerPoint._Slide slide, ref int itr, DataTable dt)
        {
            int x = 50;
            int y = 270;
            int w = 150;
            int h = 50;

            string s1, s2, s3;

            for (; itr < dt.Rows.Count; itr++)
            {
                var objRectangle = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, w, h);
                s1 = dt.Rows[itr].ItemArray[0].ToString();
                s2 = dt.Rows[itr].ItemArray[1].ToString();
                s3 = dt.Rows[itr].ItemArray[2].ToString();

                objRectangle.TextFrame.TextRange.Font.Size = 1;
                objRectangle.TextFrame.TextRange.Text = s1 + "\n" + s2 + "\n" + s3;

                if (dt.Rows[itr].ItemArray[0].ToString().Contains("GS") || dt.Rows[itr].ItemArray[0].ToString().Contains("NH"))
                    objRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGreen);
                else
                    objRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);

                objRectangle.TextFrame.TextRange.Font.Name = "Arial Narrow";
                objRectangle.TextFrame.TextRange.Paragraphs(1).Lines(1, 2).Font.Bold = MsoTriState.msoTrue;
                objRectangle.TextFrame.TextRange.Font.Size = 10;
                objRectangle.TextFrame.TextRange.Paragraphs(1).Lines(1).Font.Size = 11;
                objRectangle.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                objRectangle.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;

                x += 170;

                if (x > 730)
                {
                    x = 50;
                    y += 60;
                }
                
                if (((itr + 1) % MAX_UMD_PER_SLIDE == 0) && itr > 0)
                {
                    itr++;
                    break;
                }
            }
        }
        
        private void SetUpLowerExecPPTable(PowerPoint._Slide slide, ref int itr, DataTable dt)
        {
            int x = 50, y = 100, w = 150, h = 50;
            
            string s1, s2, s3, s4;
            
            for (; itr < dt.Rows.Count; itr++)
            {
                var objRectangle = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, w, h);
                s1 = dt.Rows[itr].ItemArray[0].ToString();
                s2 = dt.Rows[itr].ItemArray[1].ToString();
                s3 = dt.Rows[itr].ItemArray[2].ToString();
                s4 = dt.Rows[itr].ItemArray[5].ToString();

                objRectangle.TextFrame.TextRange.Font.Size = 1;
                objRectangle.TextFrame.TextRange.Text = s1 + "\n" + s2 + "\n" + s3 + "\n" + s4;

                if (dt.Rows[itr].ItemArray[0].ToString().Contains("GS") || dt.Rows[itr].ItemArray[0].ToString().Contains("NH"))
                    objRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGreen);
                else
                    objRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
                    
                if (!dt.Rows[itr].ItemArray[5].ToString().Contains(comboBox2.Text))
                {
                    objRectangle.Line.DashStyle = MsoLineDashStyle.msoLineDash;
                    objRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightGoldenrodYellow);
                }

                objRectangle.TextFrame.TextRange.Font.Name = "Arial Narrow";
                objRectangle.TextFrame.TextRange.Paragraphs(1).Lines(1, 2).Font.Bold = MsoTriState.msoTrue;
                objRectangle.TextFrame.TextRange.Font.Size = 10;
                objRectangle.TextFrame.TextRange.Paragraphs(1).Lines(1).Font.Size = 11;
                objRectangle.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                objRectangle.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;

                x += 170;

                if (x > 630)
                {
                    x = 50;
                    y += 60;
                }
            }
        }

        private DataTable CreateDataTableFromDGV(DataGridView datagrid)
        {
            DataTable dt = new DataTable();

            foreach (DataGridViewColumn col in datagrid.Columns)
            {
                dt.Columns.Add(col.Name);
            }

            foreach (DataGridViewRow row in datagrid.Rows)
            {
                DataRow drow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                    drow[cell.ColumnIndex] = cell.Value;

                dt.Rows.Add(drow);
            }

            return dt;
        }

        private void exportBtn_Click(object sender, EventArgs e)
        {
            ExportPPT();
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportPPT();
        }

        private void ExportPPT()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "PowerPoint (*.pptx)|*.pptx";
            dialog.Title = "Select a location to save your powerpoint file.";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                PowerPoint(dialog.FileName);
            }
            else
                return;
        }
        #endregion
        
        #region Merge Execute Table
        private void importExecute_Btn_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "XML (*.xml)|*.xml";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                DialogResult result = MessageBox.Show("Are you sure you want to MERGE the Execute Table from '" + Path.GetFileNameWithoutExtension(dialog.FileName) + "'?", 
                    "Merge Execute Tables", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (result == DialogResult.Yes)
                {
                    DataSet newSet = new DataSet();
                    try
                    {
                        newSet.ReadXmlSchema(ExcelParse.XSDPath);
                        newSet.ReadXml(dialog.FileName);

                        DataTable dt = newSet.Tables["ExecuteTable"];

                        DataTable dt_Old = ExcelParse.MainDataSet.Tables["ExecuteTable"];

                        dt_Old.Merge(dt, false);

                        RefreshData();

                        saveBtn.Enabled = true;
                        saveToolStripMenuItem.Enabled = true;
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("The XML file is corrupted, please re-import files", "Error opening file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                    return;
            }
        }
        #endregion
        
        #region Numeric Check
        private void numericCheck(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
                e.Handled = true;
        }
        
        private void totalReq_txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericCheck(sender, e);
        }
        
        private void en_txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericCheck(sender, e);
        }
        
        private void lg_txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericCheck(sender, e);
        }
        
        private void pk_txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericCheck(sender, e);
        }
        
        private void in_txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericCheck(sender, e);
        }
        
        private void fm_txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericCheck(sender, e);
        }
        
        private void pm_txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            numericCheck(sender, e);
        }
        #endregion
    }    
}
