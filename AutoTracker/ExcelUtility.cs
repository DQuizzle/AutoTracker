using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using LINQtoCSV;
using MoreLinq;
using ExcelDataReader;

namespace AutoTracker
{
    class ExcelUtility
    {
        #region Browse Method
        public string Browse_xlsx()
        {
            OpenFileDialog browse = new OpenFileDialog();
            browse.Filter = "Excel File|*.xlsx";
            if (browse.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return browse.FileName;
            }

            return "";
        }
        #endregion

        #region Generate CSV Method
        public bool GenerateCSV(string input1, string input2, string output)
        {
            if (HasWritePermissions(input1.Split('.')[0] + "test.txt") == false && output != "")
            {
                MessageBox.Show("Please load a valid file");
            }

            //Start Excel Import
            FileStream excelStream;
            try
            {
                excelStream = File.Open(input1, FileMode.Open, FileAccess.Read);
            }
            catch (System.IO.IOException)
            {
                MessageBox.Show("ERROR: " + Path.GetFileName(input1) + " is open. Please close the document and re-process", "Error: Import File Open", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelStream);
            var result = excelReader.AsDataSet();
            excelReader.Close();

            string excel2CSV = "";
            int row_no = 0; //Eliminate top rows of descriptions

            while (row_no < result.Tables[0].Rows.Count)
            {
                for (int i = 0; i < 12; i++)
                {
                    excel2CSV += result.Tables[0].Rows[row_no][i].ToString() + ",";
                }
                row_no++;
                excel2CSV = excel2CSV.Remove(excel2CSV.Length - 1);
                excel2CSV += "\r\n";
            }

            string outputExcel = "temp1.csv";
            File.WriteAllText(outputExcel, excel2CSV);

            //CSV Read Properties
            CsvFileDescription inputFileDesc = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = true,
                IgnoreUnknownColumns = true
            };
            CsvContext cc = new CsvContext();


            List<ExcelData> WBS = cc.Read<ExcelData>(outputExcel, inputFileDesc).ToList();

            FileStream UMD;
            try
            {
                UMD = File.Open(input2, FileMode.Open, FileAccess.Read);
            }
            catch (System.IO.IOException)
            {
                MessageBox.Show("ERROR: " + Path.GetFileName(input2) + " is open. Please close the document and re-process", "Error: Import File Open", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            IExcelDataReader UMDReader = ExcelReaderFactory.CreateOpenXmlReader(UMD);
            var UMD_Result = UMDReader.AsDataSet();
            UMDReader.Close();

            excel2CSV = "";
            row_no = 0;
            string[] combine = null;

            while (row_no < UMD_Result.Tables[0].Rows.Count)
            {
                if(UMD_Result.Tables[0].Rows[row_no][3].ToString().Contains("-"))
                {
                    combine = UMD_Result.Tables[0].Rows[row_no][3].ToString().Split('-');
                    UMD_Result.Tables[0].Rows[row_no][3] = combine[0] + "-" + UMD_Result.Tables[0].Rows[row_no][4].ToString() + "-" + combine[1];
                }

                for (int i = 0; i < 9; i++)
                {
                    excel2CSV += UMD_Result.Tables[0].Rows[row_no][i].ToString() + ",";
                }
                excel2CSV += UMD_Result.Tables[0].Rows[row_no][79].ToString() + ",";
                excel2CSV += UMD_Result.Tables[0].Rows[row_no][159].ToString() + ",";

                row_no++;
                excel2CSV = excel2CSV.Remove(excel2CSV.Length - 1);
                excel2CSV += "\r\n";
            }

            string output_UMD = "temp2.csv";
            File.WriteAllText(output_UMD, excel2CSV);

            var names = cc.Read<UMD>(output_UMD, inputFileDesc).ToList();

            List<ExcelData> outList = new List<ExcelData>();
            

            foreach (ExcelData item in WBS)
            {
                
                var asso =
                    from m in names
                    where m.WBS == item.WBSCode
                    select new { m.GRADE, m.NAME, m.MPCN, m.SERIES, m.LRMK, m.FUNDING, m.WBS_TITLE};

                foreach (var person in asso)
                {
                    if(person.FUNDING.ToString() == "FUNDED")
                    {
                        ExcelData newRow = new ExcelData(item.WBSCode, person.GRADE, person.SERIES, person.NAME, person.MPCN, person.LRMK, person.WBS_TITLE);
                        outList.Add(newRow);
                    }
                }
            }

            //CSV Write Description Properties
            CsvFileDescription outputFileDescription = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = true,
            };

            //CSV Write
            string root = "temp.csv";

            try
            {
                cc.Write(outList, root, outputFileDescription);
            }
            catch (System.IO.IOException)
            {
                System.Windows.Forms.MessageBox.Show("Please close Documents");
                return false;
            }

            return true;
        }
        #endregion

        private bool HasWritePermissions(string path)
        {

            try
            {
                File.Create(path).Close();
                File.Delete(path);
                return true;

            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }
    }
}
