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
    public partial class Main : Form
    {
        #region Initialize Variables
        private DataSet ds;
        bool loaded;

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

                comboBox1.SelectedIndex = 1;
                umdWBS.SelectedIndex = 1;
                dataGridView1.ClearSelection();
                
                FormsVisible(true);
                
                asuEditBox.Visible = true;
                deselectBtn.Enabled = true;
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
                    if (Convert.ToString(row.Cells["gradeDataGridViewTextBoxColumn"].Value).Contains("GS") || Convert.ToString(row.Cells["gradeDataGridViewTextBoxColumn"].Value).Contains("NH"))
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
            DialogResult result = MessageBox.Show("Are you sure you want to save ALL changes?", "Save Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (result == DialogResult.Yes)
            {
                ExcelParse.WriteXML(saveFile);
                saveBtn.Enabled = false;
                saveToolStripMenuItem.Enabled = false;
            }
            else
                return;
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
            }

            RefreshData();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "XML (*.xml)|*.xml";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
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
                }
                catch (Exception ex)
                {
                    MessageBox.Show("The XML file is corrupted, please re-import files", "Error opening file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            RefreshData();
        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveBtn.Enabled == true)
            {
                DialogResult result = MessageBox.Show("Do you want to save ALL changes before closing program?", "Save Changes", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

                if (result == DialogResult.Yes)
                {
                    ExcelParse.WriteXML(saveFile);
                    saveBtn.Enabled = false;
                    saveToolStripMenuItem.Enabled = false;
                    Close();
                }
                else if (result == DialogResult.No)
                    Close();
                else
                    return;
            }

            Close();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About openAbout = new About();
            openAbout.ShowDialog();
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to save ALL changes?", "Save Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (result == DialogResult.Yes)
            {
                ExcelParse.WriteXML(saveFile);
                saveBtn.Enabled = false;
                saveToolStripMenuItem.Enabled = false;
            }
            else
                return;
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (saveBtn.Enabled == true)
            {
                DialogResult result = MessageBox.Show("Do you want to save ALL changes before closing program?", "Save Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                if (result == DialogResult.Yes)
                {
                    ExcelParse.WriteXML(saveFile);
                    saveBtn.Enabled = false;
                    saveToolStripMenuItem.Enabled = false;
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
            string ID, name, MPCN;
            DataRow[] drExecList;

            DataTable dtExec = ExcelParse.MainDataSet.Tables["ExecuteTable"];

            foreach (DataRow dr in drSelected)
            {
                ID = dr["ID"].ToString();
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
            DataRow[] drExecList = dtExec.Select("MPCN='" + drExec["MPCN"] + "'");

            DataTable dtUMDExec = ExcelParse.MainDataSet.Tables["UMDTable"];
            DataRow[] drUMDExecList = dtUMDExec.Select("MPCN='" + drExec["MPCN"] + "'");

            foreach (DataRow drExecPerson in drExecList)
            {
                drExecPerson.Delete();
            }

            foreach (DataRow drUMDExec in drUMDExecList)
                drUMDExec["isExecuted"] = false;


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

        private void exec_AddBtn_Click(object sender, EventArgs e)
        {
            string reqProgID;

            if (dataGridView2.CurrentRow != null)
            {
                int i = dataGridView2.CurrentRow.Index;
                reqProgID = dataGridView2.Rows[i].Cells[3].Value.ToString();
            }
            else
                reqProgID = null;


            DataTable dt = ExcelParse.MainDataSet.Tables["UMDTable"];
            DataRow[] drMPCNChk = dt.Select("MPCN='" + execMPCN.Text + "'");

            if (reqProgID == null)
            {
                int i = dataGridView1.CurrentRow.Index;
                reqProgID = dataGridView1.Rows[i].Cells[8].Value.ToString();
            }

            DataTable newExecDt = ExcelParse.MainDataSet.Tables["ExecuteTable"];
            DataRow[] drExecIDChk = newExecDt.Select("MPCN='" + execMPCN.Text + "'");

            if (drExecIDChk.Length > 0)
            {
                MessageBox.Show(exec_NameBox.Text + " already assigned to a different program.", "Existing Person", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                DataRow newRow = newExecDt.NewRow();

                newRow["ID"] = Guid.NewGuid();
                newRow["PROG_ID"] = reqProgID;
                newRow["LRMK_ID"] = drMPCNChk[0].ItemArray[2];
                newRow["Grade"] = drMPCNChk[0].ItemArray[3];
                newRow["Series"] = drMPCNChk[0].ItemArray[4];
                newRow["Name"] = drMPCNChk[0].ItemArray[5];
                newRow["MPCN"] = drMPCNChk[0].ItemArray[6];

                foreach (DataRow row in drMPCNChk)
                    row["isExecuted"] = true;

                newExecDt.Rows.Add(newRow);

                saveBtn.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                updateUMDBtn.Enabled = false;
            }
        }

        private void addExecuteBtn_Click(object sender, EventArgs e)
        {
            string reqProgID;

            DataTable selectProgram = ExcelParse.MainDataSet.Tables["ProgTable"];
            DataRow[] drSelectProgram = selectProgram.Select("ID='" + umdProgs.SelectedValue.ToString() + "'");

            DataTable dt = ExcelParse.MainDataSet.Tables["UMDTable"];
            DataRow[] drMPCNChk = dt.Select("MPCN='" + mpcn_Txt.Text + "'");

            DataTable newExecDt = ExcelParse.MainDataSet.Tables["ExecuteTable"];
            DataRow[] drExecIDChk = newExecDt.Select("MPCN='" + mpcn_Txt.Text + "'");

            if (drExecIDChk.Length > 0)
            {
                MessageBox.Show(exec_NameBox.Text + " already assigned to a different program.", "Existing Person", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                DataRow newRow = newExecDt.NewRow();

                newRow["ID"] = Guid.NewGuid();
                newRow["PROG_ID"] = drSelectProgram[0].ItemArray[0];
                newRow["LRMK_ID"] = drMPCNChk[0].ItemArray[2];
                newRow["Grade"] = drMPCNChk[0].ItemArray[3];
                newRow["Series"] = drMPCNChk[0].ItemArray[4];
                newRow["Name"] = drMPCNChk[0].ItemArray[5];
                newRow["MPCN"] = drMPCNChk[0].ItemArray[6];

                foreach (DataRow row in drMPCNChk)
                    row["isExecuted"] = true;

                newExecDt.Rows.Add(newRow);

                saveBtn.Enabled = true;
                saveToolStripMenuItem.Enabled = true;
                updateUMDBtn.Enabled = false;
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
    }
}
