using Code7248.word_reader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = System.Windows.Forms.Application;

namespace MicrolokTools
{
    public partial class ToolForm : Form
    {
        public Document oDoc;
        public Microsoft.Office.Interop.Word.Application oWord;
        public string oFileType;
        public string oProgramID;
        public string[] stringSeparators;
        public Range oRange;
        public Regex oRegex;
        public string oCompleted;
        public string oNote;
        public Match oMatch;
        public Label[] oLabels;
        public int oLabelCount;
        public int oStart;
        public int oEnd;
        public int oVertical;
        public int oHorizontal;
        public System.Drawing.Point StartPoint;
        public string oSourceFile;
        public string oSourceFolder;
        public string oNewFile;
        public string DocText;
        public string oCurrent;
        public List<string> oFilter;
        public List<string> oProgram;
        public List<string> oNVProgram = new List<string>();
        public List<string> oSection;
        public List<string> oInput;
        public List<string> oOutput;

        public void NonVitalBuilder()
        {
            oStart = oProgram.FindIndex(x => x.ToUpper().Contains("GENISYS.SLAVE"));
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("NV.OUTPUT:"));
            oEnd = oProgram.FindIndex(oStart, x => x.ToUpper().Contains(";"));
            oOutput = oProgram.GetRange(oStart + 1, oEnd - oStart);
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("NV.INPUT:"));
            oEnd = oProgram.FindIndex(oStart, x => x.ToUpper().Contains(";"));
            oInput = oProgram.GetRange(oStart + 1, oEnd - oStart);
            oOutput = oOutput.Select(x => x.Trim()).ToList();
            oOutput.RemoveAll(x => x == "");
            oOutput = oOutput.Select(x => Regex.Replace(x, @"^H", "")).ToList();
            oInput = oInput.Select(x => x.Trim()).ToList();
            oInput.RemoveAll(x => x == "");
            oInput = oInput.Select(x => Regex.Replace(x, @"^H", "")).ToList();
            NVWrite(oProgram[0].Replace("ML", "NVML").Replace("MLK8", "MLK").Replace("ML8", "ML").Replace("MICROLOK", "GENISYS"),false);
            NVBlank();
            NVWrite("INTERFACE",false);
            NVBlank();
            NVWrite("COMM",false);
            NVBlank();
            NVWrite("LINK: GENISYS_LINE", false);
            NVWrite("ADJUSTABLE ENABLE: 0", false);
            NVWrite("PROTOCOL: GENISYS.SLAVE", false);
            NVWrite("ADJUSTABLE PORT: 3;", false);
            NVWrite("ADJUSTABLE STANDBY.PORT: 4;", false);
            NVWrite("ADJUSTABLE BAUD: 1200;", false);
            NVWrite("ADJUSTABLE STOPBITS: 1;", false);
            NVWrite("ADJUSTABLE PARITY: NONE;", false);
            NVWrite("ADJUSTABLE KEY.ON.DELAY: 50;", false);
            NVWrite("ADJUSTABLE KEY.OFF.DELAY: 50;", false);
            NVWrite("ADJUSTABLE CARRIER.MODE: CONSTANT;", false);
            NVWrite("ADJUSTABLE STALE.DATA.TIMEOUT: 60:SEC;", false);
            NVWrite("ADJUSTABLE POINT.POINT: 1;", false);
            NVBlank();
            NVWrite("ADDRESS: 0", false);
            NVWrite("ADJUSTABLE ENABLE: 1", false);
            NVBlank();
            NVWrite("NV.OUTPUT:", false);
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    NVWrite(oLine);
                }
                else
                {
                    NVWrite("HG" + oLine);
                }
            }
            NVBlank();
            NVWrite("NV.INPUT:", false);
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    NVWrite(oLine);
                }
                else
                {
                    NVWrite("HG" + oLine);
                }
            }
            NVBlank();
            NVWrite("LINK: SCS128_LINE", false);
            NVWrite("ADJUSTABLE ENABLE: 0", false);
            NVWrite("PROTOCOL: SCS.SLAVE", false);
            NVWrite("ADJUSTABLE PORT: 3;", false);
            NVWrite("ADJUSTABLE BAUD: 1200;", false);
            NVWrite("ADJUSTABLE STANDBY.PORT: 4;", false);
            NVWrite("ADJUSTABLE ALTERNATE.BAUD: 1200;", false);
            NVWrite("ADJUSTABLE STOPBITS: 1;", false);
            NVWrite("ADJUSTABLE PARITY: EVEN;", false);
            NVWrite("ADJUSTABLE KEY.ON.DELAY: 50;", false);
            NVWrite("ADJUSTABLE KEY.OFF.DELAY: 50;", false);
            NVWrite("ADJUSTABLE STALE.DATA.TIMEOUT: 60:SEC;", false);
            NVWrite("ADJUSTABLE INTERBYTE.TIMEOUT: 0:MSEC;", false);
            NVWrite("ADJUSTABLE INDICATION.ACK: ENABLED;", false);
            NVWrite("ADJUSTABLE POINT.POINT: 1;", false);
            NVBlank();
            NVWrite("ADDRESS: 0", false);
            NVWrite("ADJUSTABLE ENABLE: 1", false);
            NVBlank();
            NVWrite("NV.OUTPUT:", false);
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    NVWrite(oLine);
                }
                else
                {
                    NVWrite("HS" + oLine);
                }
            }
            NVBlank();
            NVWrite("NV.INPUT:", false);
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    NVWrite(oLine);
                }
                else
                {
                    NVWrite("HS" + oLine);
                }
            }
            NVBlank();
            NVWrite("LINK: ATCS_LINE", false);
            NVWrite("ADJUSTABLE ENABLE: 0", false);
            NVWrite("PROTOCOL: ATCS.SLAVE", false);
            NVWrite("ADJUSTABLE PORT: 3;", false);
            NVWrite("ADJUSTABLE BAUD: 9600;", false);
            NVWrite("ADJUSTABLE POLLING.TIMEOUT: 500:MSEC;", false);
            NVWrite("ADJUSTABLE POLLING.INTERVAL: 1000:MSEC;", false);
            NVWrite("ADJUSTABLE HDLC.FAIL.TIMEOUT: 60:SEC;", false);
            NVWrite("ADJUSTABLE STALE.DATA.TIMEOUT: 120:SEC;", false);
            NVWrite("ADJUSTABLE XMIT.ACK.TIMEOUT:   120:SEC;", false);
            NVWrite("ADJUSTABLE INDICATION.BROADCAST.INTERVAL: 60:SEC;", false);
            NVWrite("ADJUSTABLE MCP.ATCS.ADDRESS: \"78A2AAAAAAA1A1\";", false);
            NVWrite("ADJUSTABLE DEFAULT.ATCS.HOST.ADDRESS: \"28A2AAAAAA\";", false);
            NVWrite("ADJUSTABLE HEALTH.ATCS.ADDRESS: \"AAAAAAAAAA\";", false);
            NVWrite("ADJUSTABLE ADDRESS: \"78A2AAAAAAA2A2\"", false);
            NVBlank();
            NVWrite("ADJUSTABLE ENABLE: 1", false);
            NVWrite("STATION.NAME: XOVR;", false);
            NVBlank();
            NVWrite("NV.OUTPUT:", false);
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    NVWrite(oLine);
                }
                else
                {
                    NVWrite("HAT" + oLine);
                }
            }
            NVBlank();
            NVWrite("NV.INPUT:", false);
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    NVWrite(oLine);
                }
                else
                {
                    NVWrite("HAT" + oLine);
                }
            }
            NVBlank();
            NVWrite("LINK: UCE_LINE", false);
            NVWrite("ADJUSTABLE ENABLE: 0", false);
            NVWrite("PROTOCOL: UCE.SLAVE", false);
            NVWrite("ADJUSTABLE POINT.POINT: 1;", false);
            NVWrite("ADJUSTABLE PORT: 3;", false);
            NVWrite("ADJUSTABLE BAUD: 2400;", false);
            NVWrite("ADJUSTABLE STOPBITS: 1;", false);
            NVWrite("ADJUSTABLE PARITY: NONE;", false);
            NVWrite("ADJUSTABLE ACK.TIMEOUT: 1:SEC;", false);
            NVWrite("ADJUSTABLE XMIT.RETRY.LIMIT: 3;", false);
            NVWrite("ADJUSTABLE STALE.DATA.TIMEOUT: 60:SEC;", false);
            NVWrite("ADJUSTABLE BROADCAST.INTERVAL: 60:SEC;", false);
            NVWrite("ADJUSTABLE BUSY.TIMEOUT: 60:SEC;", false);
            NVBlank();
            NVWrite("ADDRESS: 0", false);
            NVWrite("ADJUSTABLE ENABLE: 1", false);
            NVBlank();
            NVWrite("NV.OUTPUT:", false);
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    NVWrite(oLine);
                }
                else
                {
                    NVWrite("HU" + oLine);
                }
            }
            NVBlank();
            NVWrite("NV.INPUT:", false);
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    NVWrite(oLine);
                }
                else
                {
                    NVWrite("HU" + oLine);
                }
            }
            NVBlank();
            NVWrite("LINK: ML2", false);
            NVWrite("ADJUSTABLE ENABLE: 1", false);
            NVWrite("PROTOCOL: GENISYS.MASTER", false);
            NVWrite("ADJUSTABLE PORT: 1;", false);
            NVWrite("ADJUSTABLE BAUD: 1200;", false);
            NVWrite("ADJUSTABLE STOPBITS: 1;", false);
            NVWrite("ADJUSTABLE PARITY: NONE;", false);
            NVWrite("ADJUSTABLE KEY.ON.DELAY: 12;", false);
            NVWrite("ADJUSTABLE KEY.OFF.DELAY: 12;", false);
            NVWrite("FIXED SECURE.MODE: ON;", false);
            NVWrite("FIXED MASTER.CHECKBACK: ON;", false);
            NVWrite("ADJUSTABLE POINT.POINT: 1;", false);
            NVBlank();
            NVWrite("ADDRESS: 1", false);
            NVWrite("ADJUSTABLE ENABLE: 1", false);
            NVBlank();
            NVWrite("NV.INPUT:", false);
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    NVWrite(oLine);
                }
                else
                {
                    NVWrite("L" + oLine);
                }
            }
            NVBlank();
            NVWrite("NV.OUTPUT:", false);
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    NVWrite(oLine);
                }
                else
                {
                    NVWrite("L" + oLine);
                }
            }
            NVBlank();
            NVWrite("NV.BOOLEAN BITS", false);
            NVBlank();
            NVWrite("DLVY;");
            NVBlank();
            NVWrite("TIMER BITS", false);
            NVBlank();
            NVWrite("FIXED DLVY: SET = 0:SEC CLEAR = 1:SEC;");
            NVBlank();
            NVWrite("CONFIGURATION", false);
            NVBlank();
            NVWrite("SYSTEM", false);
            NVBlank();
            NVWrite("ADJUSTABLE DEBUG_PORT_ADDRESS:      1;", false);
            NVWrite("ADJUSTABLE DEBUG_PORT_BAUDRATE:     9600;", false);
            NVWrite("LOGIC_TIMEOUT: 500:MSEC;", false);
            NVBlank();
            NVWrite("LOGIC BEGIN", false);
            NVBlank();
            NVWrite("NV.ASSIGN DLVY TO LDLVY;");
            NVBlank();
            foreach (string oLine in oInput.Where(x => x.Contains("NWZ")|| x.Contains("RWZ")))
            {
                NVWrite(String.Format("NV.ASSIGN HG{0} + HS{0} + HAT{0} + HU{0} TO L{0};", oLine.Replace(";", "").Replace(",", "")));
            }
            NVBlank();
            foreach (string oLine in oInput.Where(x => x.Contains("BLZ")))
            {
                NVWrite(String.Format("NV.ASSIGN HG{0} + HS{0} + HAT{0} + HU{0} TO L{0};", oLine.Replace(";", "").Replace(",", "")));
            }
            NVBlank();

            using (Form form = new SlotOffSelect())
            {
                form.TopLevel = true;
                form.TopMost = true;
                oVertical = 13;
                oHorizontal = 13;
                oLabelCount = (from string word in oInput where word.Contains("GZ") select word).Count();
                oLabels = new Label[oLabelCount];
                oLabelCount = 0;
                foreach (string oLine in oInput.Where(x => x.Contains("GZ")))
                {
                    oLabels[oLabelCount] = new Label();
                    form.Controls.Add(oLabels[oLabelCount]);
                    oLabels[oLabelCount].AutoSize = true;
                    oLabels[oLabelCount].Text = oLine;
                    oLabels[oLabelCount].Location = new System.Drawing.Point(oHorizontal,oVertical);
                    
                    oHorizontal = oHorizontal + oLabels[oLabelCount].Width + 13;
                    oLabelCount++;
                }
                form.ShowDialog();
            }
                












            oNewFile = oNVProgram[0].Replace("GENISYS_II PROGRAM ", "").Replace(";", "");
            oCompleted = string.Join(Environment.NewLine, oNVProgram.ToArray());
        }
        public void NVBlank()
        {
            oNVProgram.Add("");
        }
        public void NVWrite(string sString,bool oIndent = true)
        {
            if (oIndent == true)
            {
                oNVProgram.Add("  " + sString);
            }
            else
            {
                oNVProgram.Add(sString);
            }
        }
        public void WriteNonVital()
        {
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Add();
            oDoc.Content.Text = oCompleted;
            oDoc.Content.set_Style("No Spacing");
            oDoc.Content.Font.Size = 10;
            oDoc.Content.Font.Name = "courier new";
            oDoc.SaveAs2(oSourceFolder + @"\" + oNewFile + ".docx");
            oWord.Visible = true;
            //oWord.Run("Exterior.ExteriorRun");
            oDoc = null;
            oWord = null;
        }
        public void MllStripper()
        {
            oProgram = File.ReadAllLines(oSourceFile).ToList();
            oRegex = new Regex(@"-\d+-");
            oProgram.RemoveAll(x => x.Contains("\f") || x.Contains("Ansaldo STS") || x.Contains("Application:") || x.Contains("CRC =") || x.Trim() == "" || oRegex.Match(x).Success);
            //Filter compiler output
            oProgram = oGetRange(oProgram, "MICROLOK_II PROGRAM", "END PROGRAM",true);
            oRegex = new Regex(@"^( ){0,4}(\d){0,4}( ){4}");
            while (oProgram.FindIndex(x => x.Contains("//") && !oRegex.Match(x).Success) > -1)
            {
                oStart = oProgram.FindIndex(x => x.Contains("//") && !oRegex.Match(x).Success);
                oProgram[oStart] = oProgram[oStart].Replace(@"//", "  ");
            }
            oProgram = oProgram.Select(x => Regex.Replace(x, @"( ){3}\d{1}( ){4}|( ){2}\d{2}( ){4}|( ){1}\d{3}( ){4}|\d{4}( ){4}", "")).ToList();
            GetStartEnd(oProgram, "LOG BITS", ";");
            for (int i = oEnd; i > oStart + 1; i--)
            {
                oProgram[i - 1] = oProgram[i - 1] + oProgram[i];
                oProgram.RemoveAt(i);
            }

            while (oProgram.FindIndex(x => x.Contains("ASSIGN") && !x.Contains(";")) > -1)
            {
                oStart = oProgram.FindIndex(x => x.Contains("ASSIGN") && !x.Contains(";"));
                oProgram[oStart] = oProgram[oStart] + oProgram[oStart + 1];
                oProgram.RemoveAt(oStart + 1);
            }
            oCompleted = string.Join(Environment.NewLine, oProgram.ToArray());
        }
        public void GetStartEnd(List<string> sString, string sStart, string sEnd)
        {
            oStart = sString.FindIndex(x => x.ToUpper().Contains(sStart));
            oEnd = sString.FindIndex(oStart, x => x.ToUpper().Contains(sEnd));
        }
        public void RemoveNotes()
        {
            while (oProgram.FindIndex(x => x.Contains("/*")) != -1)
            {
                oStart = oProgram.FindIndex(x => x.Contains("/*"));
                oEnd = oProgram.FindIndex(oStart, x => x.Contains("*/"));
                if (oStart == oEnd)
                {
                    oProgram[oStart] = Regex.Replace(oProgram[oStart], @"(/\*).*(\*/)", "");
                }
                else
                {
                    //Clean First Line
                    oProgram[oStart] = Regex.Replace(oProgram[oStart], @"(/\*).*", "");
                    //Clean Last Line
                    oProgram[oEnd] = Regex.Replace(oProgram[oEnd], @".*\*/", "");
                    //Clear Middle Lines
                    oProgram.RemoveRange(oStart + 1, oEnd - oStart - 1);
                }
            }
            while (oProgram.FindIndex(x => x.Contains("//")) != -1)
            {
                oStart = oProgram.FindIndex(x => x.Contains("//"));
                oProgram[oStart] = Regex.Replace(oProgram[oStart], @"(//).*", "");
            }
        }
        public void CleanUp()
        {
            RemoveNotes();
            oProgram = oProgram.Select(x => x.Trim()).ToList();
            oProgram = oProgram.Where(x => x.Trim() != "").ToList();
        }
        public bool ProgramSelect()
        {
            OpenFileDialog oProgramSelect = new OpenFileDialog();
            oProgramSelect.Filter = "Microlok Files|*" + oFileType + "*";
            oProgramSelect.Title = "Select a Microlok File";
            oProgramSelect.Multiselect = false;
            if (oProgramSelect.ShowDialog() == DialogResult.OK)
            {
                oSourceFile = oProgramSelect.FileName;
                oSourceFolder = Path.GetDirectoryName(oSourceFile);
                if (oIsFileLocked() == true)
                {
                    MessageBox.Show("File is in use. Please close and try again.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return false;
            }
        }
        public bool oIsFileLocked()
        {
            FileInfo oFileInfo = new FileInfo(oSourceFile);
            FileStream stream = null;
            try
            {
                stream = oFileInfo.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }
        public static string ReplaceLast(string Source, string Find, string Replace)
        {
            int place = Source.LastIndexOf(Find);
            if (place == -1)
                return Source;
            string result = Source.Remove(place, Find.Length).Insert(place, Replace);
            return result;
        }
        public void LogBitsBuilder()
        {
            oProgram = oGetRange(oProgram, "LOCAL", "TIMER");
            oProgram = oListReplace(oProgram, ";", ",");
            oFilter = new List<string> {":","CABOUT","HEALTH","COGK","NVFLA","GE,","DLVY","FLE","PZ","BEEP","SPARE","ADJUSTABLE"};
            oProgram.RemoveAll(a => oFilter.Any(b => a.Contains(b)));
            oProgram.RemoveAll(x => !x.Contains(","));
            oProgram.RemoveAll(x => x.Contains("P") && x.Contains("K,"));
            oProgram.RemoveAll(x => x.Contains("P") && x.Contains("K1,"));
            oProgram.RemoveAll(x => x.Contains("P") && x.Contains("K2,"));
            oProgram.RemoveAll(x => x.Contains("P") && x.Contains("AS,"));
            oProgram.RemoveAll(x => x.Contains("H") && x.Contains("K,"));
            oProgram.RemoveAll(x => x.Contains("TER") && x.Contains("K"));
            oProgram.RemoveAll(x => x.Contains("LOP") && !x.Contains("LOPI"));
            oCompleted = string.Join(" ", oProgram.ToArray());
            oCompleted = ReplaceLast(oCompleted, ",", ";");
            oCompleted = "LOG BITS" + Environment.NewLine + oCompleted + Environment.NewLine;
            WriteLog();
        }
        public List<string> oListReplace(List<string> sString, string sFind, string sReplace)
        {
            sString = sString.Select(x => x.Replace(sFind, sReplace)).ToList();
            return sString;
        }
        public List<string> oGetRange(List<string> sString, string sFirst, string sSecond, bool oInclude = false)
        {
            int iFirst = sString.FindIndex(x => x.ToUpper().Contains(sFirst));
            int iSecond = sString.FindIndex(iFirst, x => x.ToUpper().Contains(sSecond));
            if(oInclude == true)
            {
                sString = sString.GetRange(iFirst, iSecond - iFirst + 1);
            }
            else
            {
                sString = sString.GetRange(iFirst + 1, iSecond - iFirst);
            }
            return sString;
        }
        public void WriteFromMll()
        {
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Add();
            oDoc.Content.Text = oCompleted;
            oDoc.Content.set_Style("No Spacing");
            oDoc.Content.Font.Size = 10;
            oDoc.Content.Font.Name = "courier new";
            oDoc.SaveAs2(oSourceFolder + @"\" + Path.GetFileNameWithoutExtension(oSourceFile) + ".docx");
            oWord.Run("Exterior.ExteriorRun");
            oCloseDoc();
        }
        public void WriteLog()
        {
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Open(oSourceFile);
            oStart = oDoc.Content.Text.IndexOf("LOG");
            oEnd = oDoc.Content.Text.IndexOf("CONFIGURATION")-1;
            oRange = oDoc.Range(Start: oStart, End: oEnd);
            oRange.Text = oCompleted;
            oDoc.Save();
            oWord.Run("Exterior.ExteriorRun");
            oCloseDoc();
        }
        public void oCloseDoc()
        {
            oDoc.Close(false);
            oWord.Visible = true;
            oWord.Activate();
            oDoc = null;
            oWord = null;
        }
        public void DocToPlain()
        {
            TextExtractor extractor = new TextExtractor(oSourceFile.ToLower());
            DocText = extractor.ExtractText();
            switch (Path.GetExtension(oSourceFile).ToUpper())
            {
                case ".DOCX":
                    stringSeparators = new string[] { "\r\n" };
                    oProgram = DocText.Split(stringSeparators, StringSplitOptions.None).ToList();
                    oProgram = oProgram.Select(s => s.Replace("\t", " ")).ToList();
                    break;
                case ".DOC":
                    stringSeparators = new string[] { "\r" };
                    oProgram = DocText.Split(stringSeparators, StringSplitOptions.None).ToList();
                    break;
                default:
                    break;
            }
            CleanUp();
        }
        public void oShow()
        {
            Show();
            CenterToScreen();
            this.TopLevel = true;
            this.TopMost = true;
        }
        private void ToolForm_Load(object sender, EventArgs e)
        {

        }
        public void LogBitsButton_Click(object sender, EventArgs e)
        {
            Hide();
            oFileType = ".doc";
            if (ProgramSelect() == true)
            {
                DocToPlain();
                LogBitsBuilder();
            }
            oShow();
        }
        public void ExtensionButton_Click(object sender, EventArgs e)
        {
            Hide();
            oFileType = ".doc";
            ProgramSelect();
            oShow();
        }
        public void NonVitalButton_Click(object sender, EventArgs e)
        {
            Hide();
            oFileType = ".doc";
            if (ProgramSelect() == true)
            {
                DocToPlain();
                NonVitalBuilder();
                WriteNonVital();
            }
            oShow();
        }
        public void MLLConvertButton_Click(object sender, EventArgs e)
        {
            Hide();
            oFileType = ".mll";
            if (ProgramSelect() == true)
            {
                File.Copy(oSourceFile, oSourceFolder + @"\" + Path.GetFileNameWithoutExtension(oSourceFile) + "-Backup.mll",true);
                MllStripper();
                WriteFromMll();
            }
            oShow();
        }
        public void oCancelButton_Click(object sender, EventArgs e)
        {
            Dispose();
        }
        public ToolForm()
        {
            InitializeComponent();
        }
    }
}
