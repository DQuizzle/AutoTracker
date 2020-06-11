using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Data;
using System.Windows.Forms;
using AutoTracker.Properties;
using AutoTracker.Tables;
using Microsoft.VisualBasic.FileIO;

using System.Diagnostics;

namespace AutoTracker
{
    class ExcelParse
    {
        #region Table Properties
        private static List<ASUTable> ASUs = new List<ASUTable>();
        private static List<ProgTable> prog_Tables = new List<ProgTable>();
        private static List<UMDTable> UMDs = new List<UMDTable>();
        private static List<ReqsTable> reqsTable = new List<ReqsTable>();
        private static DataSet mainDataSet;
        #endregion

        #region Initialize Variables
        public static DataSet MainDataSet
        {
            get { return mainDataSet; }
            set { mainDataSet = value; }
        }

        public static string XSDPath
        {
            get
            {
                return Settings.Default.XSDFileName;
            }
        }

        public static string XMLPath { get; set; }
        #endregion

        #region Execute
        public static bool Parser(string inputFile)
        {
            using (StreamReader reader = new StreamReader(inputFile))
            {
                if (File.Exists(inputFile))
                {
                    try
                    {
                        ParseCSV(inputFile);
                        BuildDataSet();
                    }
                    catch
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        public static bool WriteXML(string path)
        {
            if (mainDataSet == null)
                return false;

            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    mainDataSet.WriteXml(fs, XmlWriteMode.DiffGram);

                    return true;
                }
            }
            catch { return false; }
        }

        public static string GetFileNameOnly(string filePath)
        {
            string result = null;
            FileInfo fileInfo = new FileInfo(filePath);
            result = fileInfo.Name;
            return result;
        }

        public static void Reset()
        {
            ASUs = new List<ASUTable>();
            prog_Tables = new List<ProgTable>();
            UMDs = new List<UMDTable>();
            reqsTable = new List<ReqsTable>();
        }
        #endregion

        #region Build Data Set Method
        public static void BuildDataSet()
        {
            mainDataSet = new DataSet();
            mainDataSet.ReadXmlSchema(XSDPath);

            DataTable asutables = mainDataSet.Tables["ASUTable"];
            DataTable progTables = mainDataSet.Tables["ProgTable"];
            DataTable umdtables = mainDataSet.Tables["UMDTable"];
            DataTable reqstables = mainDataSet.Tables["ReqsTable"];

            
            foreach (var asu in ASUs)
                asutables.Rows.Add(new object[] { asu.ID, asu.WBS });

            foreach (var prog in prog_Tables)
                progTables.Rows.Add(new object[] { prog.ID, prog.WBS_ID, prog.ProgramTitle });

            foreach (var umd in UMDs)
                umdtables.Rows.Add(new object[] { umd.ID, umd.mapped_ASUID, umd.LRMK, umd.Grade, umd.Series, umd.Name, umd.MPCN });

            foreach (var req in reqsTable)
                reqstables.Rows.Add(new object[] { req.ID, req.mapped_progID, req.progName, req.LRMK, req.name, req.totalReqs, req.EN, req.LG, req.PK, req.IN, req.FM, req.PM });
        }
        #endregion

        #region Parse Method
        private static void ParseCSV(string filePath)
        { 
            DataTable csvTable = new DataTable();

            using (var csvASUFile = new TextFieldParser("temp1.csv"))
            {
                //Creates columns for all the data needed from ASU excel sheet
                csvTable.Columns.Add("WBS Code", typeof(string));
                csvTable.Columns.Add("Program Title", typeof(string));
                csvTable.Columns.Add("LRMK", typeof(string));
                csvTable.Columns.Add("Source/DELTA", typeof(string));
                csvTable.Columns.Add("Total Reqs", typeof(string));
                csvTable.Columns.Add("EN", typeof(string));
                csvTable.Columns.Add("LG", typeof(string));
                csvTable.Columns.Add("PK", typeof(string));
                csvTable.Columns.Add("IN", typeof(string));
                csvTable.Columns.Add("FM", typeof(string));
                csvTable.Columns.Add("PM", typeof(string));

                Dictionary<string, int> fieldPositions = new Dictionary<string, int>();

                csvASUFile.TextFieldType = FieldType.Delimited;
                csvASUFile.SetDelimiters(",");


                while (!csvASUFile.EndOfData)
                {
                    string[] fieldArray;

                    try
                    {
                        fieldArray = csvASUFile.ReadFields();
                    }
                    catch (MalformedLineException)
                    {
                        continue;
                    }

                    //If CSV file has a header, it grabs the count of the columns
                    if (fieldArray[0] == "WBS Code")
                    {
                        foreach (string header in fieldArray)
                        {
                            if (csvTable.Columns.Contains(header))
                            {
                                fieldPositions.Add(header, Array.IndexOf(fieldArray, header));
                            }
                        }
                    }
                    else if (fieldPositions.Count() > 0 && !string.IsNullOrEmpty(fieldArray[4]))
                    {
                        Guid matchID = new Guid();

                        //Find WBS number
                        string WBSnum = fieldArray[0].Trim();
                        ASUTable asuID = ASUs.FindLast(x => x.WBS == WBSnum);
                        if (asuID == null)
                        {
                            asuID = new ASUTable(WBSnum);
                            ASUs.Add(asuID);
                        }

                        string progTitle = fieldArray[1].Trim();
                        string LRMK = fieldArray[2].Trim(); 
                        ProgTable asuMatch = prog_Tables.FindLast(x => x.ProgramTitle != progTitle);
                        ProgTable asuProgMatch = prog_Tables.FindLast(x => x.ProgramTitle == progTitle);

                        string DELTAname = fieldArray[4].Trim();
                        string totalReqs = fieldArray[5].Trim();
                        string en = fieldArray[6].Trim();
                        string lg = fieldArray[7].Trim();
                        string pk = fieldArray[8].Trim();
                        string In = fieldArray[9].Trim();
                        string fm = fieldArray[10].Trim();
                        string pm = fieldArray[11].Trim();

                        //Ensure that the value TotalReqs exist for it to populate the rest of the data
                        if (totalReqs != "")
                        {
                            //Create a table for each unique program and assign to matching WBS code
                            if (asuProgMatch == null && progTitle != "")
                            {
                                asuProgMatch = new ProgTable(WBSnum, progTitle);
                                prog_Tables.Add(asuProgMatch);

                                matchID = asuProgMatch.ID;
                            }
                            else
                            {
                                //Each program has a possibility to have multiple lines of Req info
                                //Create a table in the ReqTable for each of those lines
                                progTitle = asuMatch.ProgramTitle;

                                //Maps the ID so they all share the same Project ID
                                ReqsTable LRMK_Match = reqsTable.FindLast(x => x.LRMK != LRMK);
                                LRMK = LRMK_Match.LRMK;
                                matchID = LRMK_Match.mapped_progID;
                            }

                            ReqsTable reqMap = new ReqsTable(matchID, progTitle, LRMK, DELTAname, totalReqs, en, lg, pk, In, fm, pm);
                            reqsTable.Add(reqMap);
                        }
                    }
                }
            }
            
            using (var csvInputFile = new TextFieldParser(filePath))
            {
                //Creates columns for all the data needed from generated csv sheet
                csvTable.Columns.Add("WBS", typeof(string));
                csvTable.Columns.Add("GRADE", typeof(string));
                csvTable.Columns.Add("OCC", typeof(string));
                csvTable.Columns.Add("NAME", typeof(string));
                csvTable.Columns.Add("MPCN", typeof(string));
                csvTable.Columns.Add("LRMK1", typeof(string));
                csvTable.Columns.Add("WBS_TITLE", typeof(string));

                Dictionary<string, int> fieldPositions = new Dictionary<string, int>();

                csvInputFile.TextFieldType = FieldType.Delimited;
                csvInputFile.SetDelimiters(",");


                while (!csvInputFile.EndOfData)
                {
                    string[] fieldArray;

                    try
                    {
                        fieldArray = csvInputFile.ReadFields();
                    }
                    catch (MalformedLineException)
                    {
                        continue;
                    }

                    if (fieldArray[0] == "WBS")
                    {
                        foreach (string header in fieldArray)
                        {
                            if (csvTable.Columns.Contains(header))
                            {
                                fieldPositions.Add(header, Array.IndexOf(fieldArray, header));
                            }
                        }
                    }
                    else if (fieldPositions.Count() > 0 && !string.IsNullOrEmpty(fieldArray[1]))
                    {
                        //Find WBS number
                        Guid progID = new Guid();

                        string wbs = fieldArray[0].Trim();
                        string grade = fieldArray[1].Trim();
                        string series = fieldArray[2].Trim();
                        string name = fieldArray[3].Trim();
                        string mpcn = fieldArray[4].Trim();
                        string LRMK = fieldArray[5].Trim();
                        string wbs_title = fieldArray[6].Trim();


                        //Each person has a unique MPCN number
                        //Check to ensure that the MPCN number does not yet exist
                        UMDTable mpcnChk = UMDs.FindLast(x => x.MPCN == mpcn);
                        if (mpcnChk == null)
                        {
                            //Second, check to see if there is a matching WBS Code
                            ProgTable progMatch = prog_Tables.Find(x => x.WBS_ID == wbs);

                            if (progMatch != null)
                            {
                                //If there is a matching WBS Code
                                //Set the Program ID from the ReqTable to the
                                //progID variable
                                progID = progMatch.ID;

                                wbs_title = CheckProgramName(wbs_title);

                                //Check LRMK matches first, if not present, go to WBS Name
                                ReqsTable lrmkMatch = reqsTable.Find(x => x.LRMK == LRMK);
                                ProgTable prNameMatch = prog_Tables.Find(x => x.ProgramTitle.ToLower().Contains(wbs_title.ToLower()));

                                //Check both the ReqsTable and the ProgTables to find a matching Project ID
                                if (lrmkMatch != null)
                                {
                                    progID = lrmkMatch.mapped_progID;
                                }
                                else if (prNameMatch != null)
                                {
                                    progID = prNameMatch.ID;
                                }
                                else
                                {
                                    string progName;
                                    if (LRMK == "")
                                        progName = wbs_title;
                                    else
                                        progName = LRMK;

                                    ProgTable newProg = new ProgTable(wbs, progName);
                                    prog_Tables.Add(newProg);

                                    progID = newProg.ID;
                                }
                            } 
                            else
                            {
                                ASUTable asuID = ASUs.FindLast(x => x.WBS == wbs);
                                if (asuID == null)
                                {
                                    asuID = new ASUTable(wbs);
                                    ASUs.Add(asuID);
                                }

                                ProgTable newProg = new ProgTable(wbs, wbs_title);
                                prog_Tables.Add(newProg);

                                progID = newProg.ID;
                            }
                            
                            UMDTable umdWBS = new UMDTable(progID, LRMK, grade, series, name, mpcn);
                            UMDs.Add(umdWBS);
                        }
                    }
                }
            }
        }
        #endregion

        #region Abbrivations
        //Check abbrivations that did not match between sheets
        private static string CheckProgramName(string progName)
        {
            if (progName == "ATP-MP")
                return "Advanced Targeting Pod Modernization Program";
            else if (progName == "ATP-SE")
                return "advanced targeting pod - sensor enhancement";
            else if (progName == "JTIDS")
                return "joint tactical information distribution system";
            else if (progName == "ALE-47 Chaff & Flare")
                return "ale-47 countermeasures dispenser system (chaff & flare)";
            else if (progName == "EGI-M")
                return "Embedded Global Positioning System/Inertial Navigation Systems-Modernization";
            else
                return progName;
        }
        #endregion
    }
        
}
