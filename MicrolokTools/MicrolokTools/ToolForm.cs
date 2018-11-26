using Code7248.word_reader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace MicrolokTools
{
    public partial class ToolForm : Form
    {
        public Document oDoc;
        public Microsoft.Office.Interop.Word.Application oWord;
        public string oFileType;
        public string[] stringSeparators;
        public Range oRange;
        public Regex oRegex;
        public string oCompleted;
        public string oNote;
        public Match oMatch;
        public int oStart;
        public int oEnd;
        public string oSourceFile;
        public string oSourceFolder;
        public string DocText;
        public List<string> oFilter;
        public List<string> oProgram;
        public string oTempFile = "TempText.txt";

        public void MllStripper()
        {
            oProgram = File.ReadAllLines(oSourceFile).ToList();
            oRegex = new Regex(@"-\d+-");
            oProgram.RemoveAll(x => x.Contains("\f") || x.Contains("Ansaldo STS") || x.Contains("Application:") || x.Contains("CRC =") || x.Trim() == "" || oRegex.Match(x).Success);
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
        public void WriteNewDoc()
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
            TextExtractor extractor = new TextExtractor(oSourceFile);
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
            ProgramSelect();
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
                WriteNewDoc();
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
