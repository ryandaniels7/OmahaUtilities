//using Code7248.word_reader;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO.Compression;

namespace MicrolokTools
{
    public class SigReq
    {
        public string GZ { get; set; }
        public string SlotOff { get; set; }
    }

    public partial class ToolForm : Form
    {
        public Document oDoc;
        public Microsoft.Office.Interop.Word.Application oWord;
        public Range oRange;
        public Regex oRegex;
        public Match oMatch;
        public MatchCollection oMatches;
        public Label oCurrentLabel;
        public Label oGZLabel;
        public Label[] oLabels;
        public ComboBox OSBox;
        public int oLabelCount;
        public int oStart;
        public int oEnd;
        public int oVertical;
        public int oHorizontal;
        public int oBoardCount;
        public int oInputIndex;
        public int oOutputIndex;
        public double oWidth;
        public Form SlotForm;
        public Form oHeaderForm;
        public Form SlotTracks;
        public Form TrackForm;
        public System.Drawing.Point StartPoint;
        public string oLinkMii;
        public string oLinkPeer;
        public string oFlashMii;
        public string oFlashPeer;
        public string oSourceFile;
        public string oNewFile;
        public string oSourceFolder;
        public string oNewFolder;
        public string DocText;
        public string oCurrent;
        public string oProgramName;
        public string oCompleted;
        public object objStart;
        public object objEnd;
        public string SomePropertyName { get; set; }
        public string[] stringSeparators;
        public List<SigReq> oSigReqs = new List<SigReq>();
        public List<SigReq> oCOSigReqs = new List<SigReq>();
        public List<string> oSlotOffs;
        public List<string> oSwitches = new List<string>();
        public List<string> oSection;
        public List<string> oSlotLabels;
        public List<string> oFilter;
        public List<string> oProgram;
        public List<string> sProgram = new List<string>();
        public List<string> oExtensions = new List<string>();
        public List<string> oGZs = new List<string>();
        //public List<string> sProgram = new List<string>();
        public List<string> oInput;
        public List<string> oOutput;
        public List<string> oBoolean;
        public List<string> oAllBoolean = new List<string>();
        public List<string> oTables = new List<string>();
        public List<string> oIOBoolean = new List<string>();
        public List<string> oVBoolean = new List<string>();
        public List<string> oNVBoolean = new List<string>();
        public List<string> oEquations;
        public List<string> oFilteredEquations;

        public void DocToPlain()
        {
            switch (Path.GetExtension(oSourceFile).ToUpper())
            {
                case ".DOCX":
                    ZipArchive oArchive = ZipFile.Open(oSourceFile, ZipArchiveMode.Read);
                    ZipArchiveEntry oEntry = oArchive.GetEntry("word/document.xml");
                    using (var oReader = new StreamReader(oEntry.Open()))
                    {
                        oCurrent = oReader.ReadToEnd();
                    }
                    oArchive.Dispose();
                    //Clipboard.SetText(oCurrent);
                    oCurrent = oCurrent.Replace("</w:p>", "</w:p>" + Environment.NewLine).Replace(@" xml:space=""preserve""", "").Replace("amp;", "");
                    MatchCollection oMatches = Regex.Matches(oCurrent, @"(?<=<w:t>)([^<]*)|" + Environment.NewLine);
                    oCurrent = "";
                    foreach (Match oMatch in oMatches)
                    {
                        oCurrent = oCurrent + oMatch;
                    }

                    stringSeparators = new string[] { Environment.NewLine };
                    oProgram = oCurrent.Split(stringSeparators, StringSplitOptions.None).ToList();
                    //oProgram = oProgram.Select(x => Regex.Match(x, @"(?<=<w:t>)([^<]*)").Value).ToList();
                    //oProgram = oProgram.Select(x => Regex.Replace(x, @"(.*)(<w:t>)", "")).ToList();
                    //oProgram = oProgram.Select(x => Regex.Replace(x, @"(</w:t>)(.*)", "")).ToList();
                    //oProgram = oProgram.Select(x => Regex.Replace(x, @"<(.*)>", "")).ToList();
                    break;
                case ".DOC":
                    oCurrent = File.ReadAllText(oSourceFile);
                    oStart = oCurrent.IndexOf("MICROLOK");
                    oEnd = oCurrent.LastIndexOf("END PROGRAM") + 11;
                    oCurrent = oCurrent.Substring(oStart, oEnd - oStart);
                    stringSeparators = new string[] { "\r" };
                    oProgram = oCurrent.Split(stringSeparators, StringSplitOptions.None).ToList();
                    break;
                default:
                    break;      
            }
            CleanUp();
        }
        public void FormatWrite(string psText, string oBold = "")
        {
            oEnd = oDoc.Characters.Count;
            Paragraph para = oDoc.Content.Paragraphs.Add();
            para.Range.Text = psText;
            // Explicitly set this to "not bold"
            para.Range.Font.Bold = 0;
            para.Range.Font.Size = 10;
            object add = oBold.Length;
            para.Range.Font.Name = "courier new";
            try
            {
                objStart = psText.IndexOf(oBold) + oEnd - 2;
                objEnd = psText.IndexOf(oBold) + oEnd - 1 + oBold.Length;
                oRange = oDoc.Range(ref objStart, ref objEnd);
            }
            catch
            {
                objStart = psText.IndexOf(oBold) + oEnd - 1;
                objEnd = psText.IndexOf(oBold) + oEnd + oBold.Length;
                oRange = oDoc.Range(ref objStart, ref objEnd);
            }
            if (!oRange.Text.Contains("~"))
            {
                objStart = psText.IndexOf(oBold) + oEnd - 1;
                oRange = oDoc.Range(ref objStart, ref objEnd);
            }
            oRange.Bold = 1;
            para.Range.InsertParagraphAfter();
        }

        public void CheckSwitches()
        {
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Add();
            oDoc.Content.set_Style("No Spacing");
            oEnd = 0;
            //oSourceFolder = Directory.GetParent(oSourceFolder) + @"\QLCP" + oProgramName;
            //Directory.CreateDirectory(oSourceFolder);
            
            oTables = oGetRange(oProgram, "LOCAL", "BOOLEAN");
            oSwitches = oTables.Where(x => x.Contains("NJPI")).ToList();
            oSwitches = oSwitches.Select(x => x.Replace("NJPI,", "").Replace("NJPI;", "")).ToList();
            oEquations = oGetRange(oProgram, "LOGIC BEGIN", "END LOGIC");
            foreach (string oSwitch in oSwitches)
            {
                if (oSwitch.Contains("RI"))
                {
                    oCurrent = oSwitch.Replace("RI", "").Replace("_", "");
                    FormatWrite(oCurrent + "NWPI", oCurrent + "NWPI");
                    foreach (string oLine in oEquations.Where(x => x.Contains(oCurrent + "NWPI")))
                    {
                        FormatWrite(oLine, oCurrent + "NWPI");
                    }
                    FormatWrite("");
                    FormatWrite(oCurrent + "RWPI", oCurrent + "RWPI");
                    foreach (string oLine in oEquations.Where(x => x.Contains(oCurrent + "RWPI")))
                    {
                        FormatWrite(oLine, oCurrent + "RWPI");
                    }
                    FormatWrite("");
                    FormatWrite(oCurrent + "NJPI", oCurrent + "NJPI");
                    foreach (string oLine in oEquations.Where(x => x.Contains(oCurrent + "NJPI")))
                    {
                        FormatWrite(oLine, oCurrent + "NJPI");
                    }
                    FormatWrite("");
                    FormatWrite(oCurrent + "NWCK", oCurrent + "NWCK");
                    foreach (string oLine in oEquations.Where(x => x.Contains(oCurrent + "NWCK")))
                    {
                        FormatWrite(oLine, oCurrent + "NWCK");
                    }
                    FormatWrite("");
                    FormatWrite(oCurrent + "RWCK", oCurrent + "RWCK");
                    foreach (string oLine in oEquations.Where(x => x.Contains(oCurrent + "RWCK")))
                    {
                        FormatWrite(oLine, oCurrent + "RWCK");
                    }
                }
                else
                {
                    FormatWrite("");
                    FormatWrite(oSwitch + "NWPI", oSwitch + "NWPI");
                    foreach (string oLine in oEquations.Where(x => x.Contains(oSwitch + "NWPI")))
                    {
                        FormatWrite(oLine, oSwitch + "NWPI");
                    }
                    FormatWrite("");
                    FormatWrite(oSwitch + "RWPI", oSwitch + "RWPI");
                    foreach (string oLine in oEquations.Where(x => x.Contains(oSwitch + "RWPI")))
                    {
                        FormatWrite(oLine, oSwitch + "RWPI");
                    }
                    FormatWrite("");
                    FormatWrite(oSwitch + "NJPI", oSwitch + "NJPI");
                    foreach (string oLine in oEquations.Where(x => x.Contains(oSwitch + "NJPI")))
                    {
                        FormatWrite(oLine, oSwitch + "NJPI");
                    }
                    FormatWrite("");
                    FormatWrite(oSwitch + "NWCK", oSwitch + "NWCK");
                    foreach (string oLine in oEquations.Where(x => x.Contains(oSwitch + "NWCK")))
                    {
                        FormatWrite(oLine, oSwitch + "NWCK");
                    }
                    FormatWrite("");
                    FormatWrite(oSwitch + "RWCK", oSwitch + "RWCK");
                    foreach (string oLine in oEquations.Where(x => x.Contains(oSwitch + "RWCK")))
                    {
                        FormatWrite(oLine, oSwitch + "RWCK");
                    }
                    FormatWrite("");
                }
            }
            oWord.Visible = true;
            //oDoc.SaveAs2(oSourceFolder + @"\Switch_" + oProgramName + ".docx");
        }

        public void CheckSwitches2()
        {
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Add();
            oDoc.Content.set_Style("No Spacing");
            oEnd = 0;
            //oSourceFolder = Directory.GetParent(oSourceFolder) + @"\QLCP" + oProgramName;
            //Directory.CreateDirectory(oSourceFolder);

            oSection = oGetRange(oProgram, "LOCAL", "BOOLEAN");
            while (oSection.FindIndex(x => x.Contains("NJPI")) > -1)
            {
                oStart = oSection.FindIndex(x => x.ToUpper().Contains("NJPI"));
                oCurrent = oSection[oStart].Replace("NJPI", "").Replace(",", "").Replace(";", "");
                oSwitches.Add(oCurrent);
                oSection[oStart] = "";
            }
            oSection = oGetRange(oProgram, "LOGIC BEGIN", "END LOGIC");
            foreach (string oSwitch in oSwitches)
            {
                if (oSwitch.Contains("RI"))
                {
                    oCurrent = oSwitch.Replace("RI", "").Replace("_", "");
                    FormatWrite(oCurrent + "NWPI", oCurrent + "NWPI");
                    foreach (string oLine in oSection.Where(x => x.Contains(oCurrent + "NWPI")))
                    {
                        FormatWrite(oLine, oCurrent + "NWPI");
                    }
                    FormatWrite("");
                    FormatWrite(oCurrent + "RWPI", oCurrent + "RWPI");
                    foreach (string oLine in oSection.Where(x => x.Contains(oCurrent + "RWPI")))
                    {
                        FormatWrite(oLine, oCurrent + "RWPI");
                    }
                    FormatWrite("");
                    FormatWrite(oCurrent + "NJPI", oCurrent + "NJPI");
                    foreach (string oLine in oSection.Where(x => x.Contains(oCurrent + "NJPI")))
                    {
                        FormatWrite(oLine, oCurrent + "NJPI");
                    }
                    FormatWrite("");
                    FormatWrite(oCurrent + "NWCK", oCurrent + "NWCK");
                    foreach (string oLine in oSection.Where(x => x.Contains(oCurrent + "NWCK")))
                    {
                        FormatWrite(oLine, oCurrent + "NWCK");
                    }
                    FormatWrite("");
                    FormatWrite(oCurrent + "RWCK", oCurrent + "RWCK");
                    foreach (string oLine in oSection.Where(x => x.Contains(oCurrent + "RWCK")))
                    {
                        FormatWrite(oLine, oCurrent + "RWCK");
                    }
                }
                else
                {
                    FormatWrite("");
                    FormatWrite(oSwitch + "NWPI", oSwitch + "NWPI");
                    foreach (string oLine in oSection.Where(x => x.Contains(oSwitch + "NWPI")))
                    {
                        FormatWrite(oLine, oSwitch + "NWPI");
                    }
                    FormatWrite("");
                    FormatWrite(oSwitch + "RWPI", oSwitch + "RWPI");
                    foreach (string oLine in oSection.Where(x => x.Contains(oSwitch + "RWPI")))
                    {
                        FormatWrite(oLine, oSwitch + "RWPI");
                    }
                    FormatWrite("");
                    FormatWrite(oSwitch + "NJPI", oSwitch + "NJPI");
                    foreach (string oLine in oSection.Where(x => x.Contains(oSwitch + "NJPI")))
                    {
                        FormatWrite(oLine, oSwitch + "NJPI");
                    }
                    FormatWrite("");
                    FormatWrite(oSwitch + "NWCK", oSwitch + "NWCK");
                    foreach (string oLine in oSection.Where(x => x.Contains(oSwitch + "NWCK")))
                    {
                        FormatWrite(oLine, oSwitch + "NWCK");
                    }
                    FormatWrite("");
                    FormatWrite(oSwitch + "RWCK", oSwitch + "RWCK");
                    foreach (string oLine in oSection.Where(x => x.Contains(oSwitch + "RWCK")))
                    {
                        FormatWrite(oLine, oSwitch + "RWCK");
                    }
                    FormatWrite("");
                }
            }
            oWord.Visible = true;
            //oDoc.SaveAs2(oSourceFolder + @"\Switch_" + oProgramName + ".docx");
        }

        public void CreateBoolean()
        {
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Add();
            oDoc.Content.set_Style("No Spacing");
            oEnd = 0;

            oTables = oGetRange(oProgram, "LOCAL", "BOOLEAN BITS");
            oBoolean = oGetRange(oProgram, "BOOLEAN", "TIMER BITS");
            oEquations = oGetRange(oProgram, "LOGIC BEGIN", "END PROGRAM");
            oRegex = new Regex(@"(\w*)(?=,|;)");
            foreach (var line in oTables)
            {
                oMatch = oRegex.Match(line);
                if (oMatch.Success && !line.Contains("SPARE") && !line.Contains(":"))
                {
                    oIOBoolean.Add(oMatch.Value);
                }
            }
            oIOBoolean.Add("KILL");
            oRegex = new Regex(@"(?<= |,)(\w*)(?=,|;)");
            foreach (var line in oEquations)
            {
                oMatches = oRegex.Matches(line);
                if (oMatches.Count > 0 && line.Contains("NV.ASSIGN"))
                {
                    foreach (Match match in oMatches)
                    {
                        oNVBoolean.Add(match.Value);
                    }
                }
                else if (oMatches.Count > 0 && line.Contains("ASSIGN"))
                {
                    foreach (Match match in oMatches)
                    {
                        oVBoolean.Add(match.Value);
                    }
                }
            }
            oVBoolean = oVBoolean.Except(oIOBoolean).ToList();
            oNVBoolean = oNVBoolean.Except(oIOBoolean).ToList();

            oVBoolean = oVBoolean.Select(x => "  " + x + ",").ToList();
            oNVBoolean = oNVBoolean.Select(x => "  " + x + ",").ToList();
            oVBoolean[oVBoolean.Count - 1] = oVBoolean[oVBoolean.Count - 1].Replace(",", ";");
            oNVBoolean[oNVBoolean.Count - 1] = oNVBoolean[oNVBoolean.Count - 1].Replace(",", ";");

            oCompleted = "BOOLEAN BITS" + Environment.NewLine + Environment.NewLine;
            oCompleted += string.Join(Environment.NewLine, oVBoolean.ToArray());
            oCompleted += Environment.NewLine + Environment.NewLine + "NV.BOOLEAN BITS" + Environment.NewLine + Environment.NewLine;
            oCompleted += string.Join(Environment.NewLine, oNVBoolean.ToArray());
            Clipboard.SetText(oCompleted);
        }
        public void ExtBuilder()
        {
            oProgramName = Regex.Match(oProgram.FirstOrDefault(s => s.Contains("PROGRAM ")), @"(?<= PROGRAM )(.*)(?=;)").Value;
            oProgramName = Regex.Replace(oProgramName, @"M(?=\d)", "E");
            using (oHeaderForm = new HeaderForm())
            {
                oHeaderForm.TopLevel = true;
                oHeaderForm.TopMost = true;
                oHeaderForm.ShowDialog();
            }
            oStart = oProgram.FindIndex(x => x.ToUpper().Contains("VTLCOMMS"));
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("STATION") && x.ToUpper().Contains("EXT"));
            oExtensions.Add(oProgram[oStart]);
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("OUTPUT:"));
            oEnd = oProgram.FindIndex(oStart, x => x.ToUpper().Contains(";"));
            oOutput = oProgram.GetRange(oStart + 1, oEnd - oStart);
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("INPUT:"));
            oEnd = oProgram.FindIndex(oStart, x => x.ToUpper().Contains(";"));
            oInput = oProgram.GetRange(oStart + 1, oEnd - oStart);
            oOutput = oOutput.Select(x => x.Trim()).ToList();
            oOutput = oOutput.Select(x => x.Replace(" ", "")).ToList();
            oOutput.RemoveAll(x => x == "");
            oInput = oInput.Select(x => x.Trim()).ToList();
            oInput = oInput.Select(x => x.Replace(" ", "")).ToList();
            oInput.RemoveAll(x => x == "");
            FormatIO();
            oWrite(String.Format("MICROLOK_II PROGRAM {0};", oProgramName));
            oBlank();
            oWrite("/*");
            oBlank();
            oWrite("XX000 - EXTENSION");
            oWrite("M.P. X.XX");
            oWrite("XXXXX SUB");
            oWrite("CHASSIS ID = XXXX");
            oBlank();
            oWrite("PROGRAM HISTORY & REVISION LEVEL");
            oBlank();
            oWrite("REV LEVEL      DESCRIPTION                             CHANGED BY");
            oWrite("NC             NEW PROGRAM                             XXX");
            oBlank();
            oWrite("*/");
            oBlank();
            oWrite("INTERFACE");
            oBlank();
            oWrite("LOCAL");
            oBlank();
            oWrite("BOARD: NVIO");
            oWrite("ADJUSTABLE ENABLE: 1");
            oWrite("TYPE: NV.IN32.OUT32");
            oBlank();
            oWrite("NV.OUTPUT:");
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SONALERT;", true);
            oBlank();
            oWrite("NV.INPUT:");
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("SPARE,", true);
            oWrite("LAMPTEST,", true);
            oWrite("PO,", true);
            oWrite("SPARE,", true);
            oWrite("SWIK,", true);
            oWrite("SPARE,", true);
            oWrite("FIBEROK;", true);
            oBlank();
            oCompleted = string.Join(Environment.NewLine, sProgram.ToArray());
            }
            public void QLCPBuilder()
        {
            int xCount = 24;
            oProgramName = Regex.Match(oProgram.FirstOrDefault(s => s.Contains("PROGRAM ")), @"(?<= PROGRAM )(.*)(?=;)").Value;
            oNewFile = "QLCP_" + oProgramName;
            oNewFolder = Directory.GetParent(oSourceFolder) + @"\" + oNewFile;
            PullAddress();
            oStart = oProgram.FindIndex(x => x.ToUpper().Contains("LINK: QLCP"));
            if (oStart == -1)
            {
                oStart = oProgram.FindIndex(x => x.ToUpper().Contains("LINK: QCOMMS"));
            }
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("NV.OUTPUT:"));
            oEnd = oProgram.FindIndex(oStart, x => x.ToUpper().Contains(";"));
            oOutput = oProgram.GetRange(oStart + 1, oEnd - oStart);
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("NV.INPUT:"));
            oEnd = oProgram.FindIndex(oStart, x => x.ToUpper().Contains(";"));
            oInput = oProgram.GetRange(oStart + 1, oEnd - oStart);
            oOutput = oOutput.Select(x => x.Trim()).ToList();
            oOutput.RemoveAll(x => x == "");
            oOutput = oOutput.Select(x => Regex.Replace(x, @"^P", "")).ToList();
            oOutput = oOutput.Select(x => Regex.Replace(x, @"^1P", "")).ToList();
            oOutput = oOutput.Select(x => x.Replace(" ","").Replace("PLOCAL", "LOCAL").Replace("PST", "ST")).ToList();
            oInput = oInput.Select(x => x.Trim()).ToList();
            oInput.RemoveAll(x => x == "");
            oInput = oInput.Select(x => Regex.Replace(x, @"^P", "")).ToList();
            oInput = oInput.Select(x => Regex.Replace(x, @"^1P", "")).ToList();
            oInput = oInput.Select(x => x.Replace("1LOCAL", "LOCAL").Replace("1ST_ENA","ST_ENA")).ToList();
            oInput = oInput.Select(x => x.Replace(" ", "")).ToList();
            FormatIO();

            oWrite(String.Format("MICROLOK_II PROGRAM QLCP_{0};", oProgramName));
            oBlank();
            oWrite("INTERFACE");
            oBlank();
            oWrite("LOCAL");
            oBlank();
            oBoardCount = (oInput.Count + 31) / 32;
            if ((oOutput.Count + 31) / 32 > oBoardCount)
            {
                oBoardCount = (oOutput.Count + 31) / 32;
            }
            for (int i = 0; i <= oBoardCount - 1; i++)
            {
                oWrite("BOARD:NVIO" + (i + 1));
                oWrite("FIXED ENABLE:1");
                oWrite("TYPE: NV.IN32.OUT32");
                oBlank();
                oWrite("NV.OUTPUT:");
                if (oOutput.Count - ((i + 1) * 32) >= 0)
                {
                    oStart = i * 32;
                    oEnd = oOutput.Count - oStart;
                    if (oEnd > 32)
                    {
                        oEnd = 32;
                    }
                    oSection = oOutput.GetRange(oStart, oEnd);
                    foreach (var oLine in oSection)
                    {
                        if (oLine.Contains("SPARE"))
                        {
                            oWrite(oLine, true);
                        }
                        else
                        {
                            oWrite("P" + oLine, true);
                        }
                    }
                }
                else
                {
                    oWrite("SPARE;");
                }
                sProgram[sProgram.Count - 1] = sProgram[sProgram.Count - 1].Replace(",", ";");
                oBlank();
                oWrite("NV.INPUT:");
                if (oInput.Count - ((i + 1) * 32) >= 0)
                {
                    oStart = i * 32;
                    oEnd = oInput.Count - oStart;
                    if (oEnd > 32)
                    {
                        oEnd = 32;
                    }
                    oSection = oInput.GetRange(oStart, oEnd);
                    foreach (var oLine in oSection)
                    {
                        if (oLine.Contains("SPARE"))
                        {
                            oWrite(oLine, true);
                        }
                        else
                        {
                            oWrite("P" + oLine, true);
                        }
                    }
                }
                else
                {
                    oWrite("SPARE;");
                }
                sProgram[sProgram.Count - 1] = sProgram[sProgram.Count - 1].Replace(",", ";");
                oBlank();
            }
            oWrite("COMM");
            oBlank();
            oWrite("LINK: QLCPCOMMS");
            oWrite("ADJUSTABLE ENABLE: 1");
            oWrite("PROTOCOL: MII.PEER");
            oWrite("ADJUSTABLE PORT: 2;");
            oWrite("ADJUSTABLE POINT.POINT: 1;");
            oWrite("ADJUSTABLE BAUD: 38400;");
            oWrite("FIXED STOPBITS: 1;");
            oWrite("FIXED PARITY: NONE;");
            oWrite("FIXED KEY.ON.DELAY: 0;");
            oWrite("FIXED KEY.OFF.DELAY: 0;");
            oWrite("ADJUSTABLE GRANT.DELAY: 1000:MSEC;");
            oWrite("FIXED MII.NV.ADDRESS: " + oLinkMii);
            oWrite("ADJUSTABLE ENABLE: 1");
            oWrite("STATION.NAME: LCP_LINK;");
            oWrite("FIXED PEER.ADDRESS: " + oLinkPeer + ";");
            oWrite("ADJUSTABLE ACK.TIMEOUT: 2000:MSEC;");
            oWrite("ADJUSTABLE HEARTBEAT.INTERVAL: 2000:MSEC;");
            oWrite("ADJUSTABLE INDICATION.UPDATE.CYCLE: 10;");
            oWrite("ADJUSTABLE STALE.DATA.TIMEOUT: 5000:MSEC;");
            oWrite("ADJUSTABLE CLOCK.MASTER: 0;");
            oBlank();
            oWrite("NV.OUTPUT:");
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else
                {
                    oWrite("P" + NoPunct(oLine) + "_OUT,", true);
                }
            }
            sProgram[sProgram.Count - 1] = sProgram[sProgram.Count - 1].Replace(",", ";");
            oBlank();
            oWrite("NV.INPUT:");
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else
                {
                    oWrite("P" + NoPunct(oLine) + "_IN,", true);
                }
            }
            sProgram[sProgram.Count - 1] = sProgram[sProgram.Count - 1].Replace(",", ";");
            oBlank();
            oWrite("FIXED MII.NV.ADDRESS: " + oFlashMii);
            oWrite("ADJUSTABLE ENABLE: 1");
            oWrite("STATION.NAME: FLASHING_QLCP;");
            oWrite("FIXED PEER.ADDRESS: " + oFlashPeer + ";");
            oWrite("ADJUSTABLE ACK.TIMEOUT: 2000:MSEC;");
            oWrite("ADJUSTABLE HEARTBEAT.INTERVAL: 2000:MSEC;");
            oWrite("ADJUSTABLE INDICATION.UPDATE.CYCLE: 10;");
            oWrite("ADJUSTABLE STALE.DATA.TIMEOUT: 5000:MSEC;");
            oWrite("ADJUSTABLE CLOCK.MASTER: 0;");
            oWrite("NV.OUTPUT:");
            oWrite("SPARE;",true);
            oBlank();
            oWrite("NV.INPUT:");
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else
                {
                    oWrite("FL_" + NoPunct(oLine) + ",", true);
                }
            }
            sProgram[sProgram.Count - 1] = sProgram[sProgram.Count - 1].Replace(",", ";");
            oBlank();
            oWrite("NV.BOOLEAN BITS");
            for (int i = 1; i <= xCount; i++)
            {
                oWrite("XX" + i + ",");
            }
            oWrite("NVFLA;");
            oBlank();
            oWrite("TIMER BITS",true);
            oWrite("FIXED NVFLA: SET = 500:MSEC CLEAR = 500:MSEC;");
            oBlank();
            oWrite("CONFIGURATION");
            oBlank();
            oWrite("SYSTEM");
            oBlank();
            oWrite("FIXED DEBUG_PORT_ADDRESS:2;",true);
            oWrite("FIXED DEBUG_PORT_BAUDRATE:9600;",true);
            oBlank();
            oWrite("USER BIT");
            oWrite("NORMAL: \"NORMAL\", 0;", true);
            oWrite("REVERSE: \"REVERSE\", 0; ", true);
            oBlank();
            oWrite("LOGIC BEGIN");
            oBlank();
            oWrite("ASSIGN 1 TO CPS.ENABLE;");
            oWrite("NV.ASSIGN 1 * ~NVFLA TO NVFLA;");
            for (int i = 1; i <= xCount; i++)
            {
                oWrite("NV.ASSIGN 0 * XX" + i + " TO XX" + i + ";");
            }
            foreach (string oLine in oOutput.Where(x => !x.Contains("SPARE")))
            {
                oWrite(String.Format("NV.ASSIGN P{0}_IN * (FL_{0} * NVFLA + ~FL_{0}) TO P{0};", oLine.Replace(";", "").Replace(",", "")));
            }
            oBlank();
            foreach (string oLine in oInput.Where(x => !x.Contains("SPARE")))
            {
                if (oLine.Contains("NWZ"))
                {
                    oWrite(String.Format("NV.ASSIGN P{0} + NORMAL TO P{0}_OUT;", oLine.Replace(";", "").Replace(",", "")));
                }
                else if (oLine.Contains("RWZ"))
                {
                    oWrite(String.Format("NV.ASSIGN P{0} + REVERSE TO P{0}_OUT;", oLine.Replace(";", "").Replace(",", "")));
                }
                else
                {
                    oWrite(String.Format("NV.ASSIGN P{0} TO P{0}_OUT;", oLine.Replace(";", "").Replace(",", "")));
                }
            }
            oBlank();
            oWrite("END LOGIC");
            oWrite("END PROGRAM");
            oCompleted = string.Join(Environment.NewLine, sProgram.ToArray());
        }
        public void PullAddress()
        {
            oStart = oProgram.FindIndex(x => x.ToUpper().Contains("LINK: QLCP"));
            if (oStart == -1)
            {
                oStart = oProgram.FindIndex(x => x.ToUpper().Contains("LINK: QCOMMS"));
            }
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("FIXED MII.NV.ADDRESS: "));
            oLinkPeer = oProgram[oStart].Replace("FIXED MII.NV.ADDRESS: ", "");
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("FIXED PEER.ADDRESS: "));
            oLinkMii = oProgram[oStart].Replace("FIXED PEER.ADDRESS: ", "").Replace(";", "");
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("FIXED MII.NV.ADDRESS: "));
            oFlashPeer = oProgram[oStart].Replace("FIXED MII.NV.ADDRESS: ", "");
            oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("FIXED PEER.ADDRESS: "));
            oFlashMii = oProgram[oStart].Replace("FIXED PEER.ADDRESS: ", "").Replace(";", "");
        }
        public void FormatIO()
        {
            oInput = oInput.Select(x => x.Replace(";", ",")).ToList();
            while(oInput.Count % 32 != 0)
            {
                oInput.Add("SPARE,");
            }
            oInput[oInput.Count-1] = oInput[oInput.Count-1].Replace(",", ";");
            oOutput = oOutput.Select(x => x.Replace(";", ",")).ToList();
            while (oOutput.Count % 32 != 0)
            {
                oOutput.Add("SPARE,");
            }
            oOutput[oOutput.Count - 1] = oOutput[oOutput.Count - 1].Replace(",", ";");
        }
        public void GrabSlotOff()
        {
            oFilteredEquations = oEquations.Where(x => x.Contains("_GZS;")).ToList();
            foreach(var equation in oFilteredEquations)
            {
                SigReq NewSigReq = new SigReq();
                oMatch = Regex.Match(equation, @"(H\S*GZ)");
                NewSigReq.GZ = oMatch.Value.ToString().Replace("H","");
                oMatch = Regex.Match(equation, @"(\S*TT)");
                NewSigReq.SlotOff = oMatch.Value.ToString().Replace("TT","TK");
                oSigReqs.Add(NewSigReq);
            }
        }
        public void GrabCOSlotOff()
        {
            oFilteredEquations = oEquations.Where(x => x.Contains("_COGZS;")).ToList();
            List<string> usedCOGZ = new List<string>();
            foreach (var equation in oFilteredEquations)
            {
                SigReq NewSigReq = new SigReq();
                oMatch = Regex.Match(equation, @"(H\S*COGZ)");
                string result = usedCOGZ.FirstOrDefault(x => x == oMatch.Value);
                if (result == null)
                {
                    usedCOGZ.Add(oMatch.Value);
                    NewSigReq.GZ = oMatch.Value.ToString().Replace("H", "");
                    oMatch = Regex.Match(equation, @"(\S*TT)");
                    NewSigReq.SlotOff = oMatch.Value.ToString().Replace("TT", "TK");
                    oCOSigReqs.Add(NewSigReq);
                    oCOSigReqs.Distinct();
                }
            }
        }
        public void NonVitalBuilder()
        {
            oEquations = oGetRange(oProgram, "LOGIC BEGIN", "END PROGRAM");
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
            oOutput = oOutput.Select(x => x.Replace(" ","")).ToList();
            oInput = oInput.Select(x => x.Trim()).ToList();
            oInput.RemoveAll(x => x == "");
            oInput = oInput.Select(x => Regex.Replace(x, @"^H", "")).ToList();
            oInput = oInput.Select(x => x.Replace(" ", "")).ToList();
            oOutput[oOutput.LastIndexOf("SPARE,")] = "BKUP,";
            oWrite(oProgram[0].Replace("ML", "NVML").Replace("MLK8", "MLK").Replace("ML8", "ML").Replace("MICROLOK", "GENISYS"));
            oBlank();
            oWrite("INTERFACE");
            oBlank();
            oWrite("COMM");
            oBlank();
            oWrite("LINK: GENISYS_LINE");
            oWrite("ADJUSTABLE ENABLE: 0");
            oWrite("PROTOCOL: GENISYS.SLAVE");
            oWrite("ADJUSTABLE PORT: 3;");
            oWrite("ADJUSTABLE STANDBY.PORT: 4;");
            oWrite("ADJUSTABLE BAUD: 1200;");
            oWrite("ADJUSTABLE STOPBITS: 1;");
            oWrite("ADJUSTABLE PARITY: NONE;");
            oWrite("ADJUSTABLE KEY.ON.DELAY: 50;");
            oWrite("ADJUSTABLE KEY.OFF.DELAY: 50;");
            oWrite("ADJUSTABLE CARRIER.MODE: CONSTANT;");
            oWrite("ADJUSTABLE STALE.DATA.TIMEOUT: 60:SEC;");
            oWrite("ADJUSTABLE POINT.POINT: 1;");
            oBlank();
            oWrite("ADDRESS: 0");
            oWrite("ADJUSTABLE ENABLE: 1");
            oBlank();
            oWrite("NV.OUTPUT:");
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else
                {
                    oWrite("HG" + oLine, true);
                }
            }
            oBlank();
            oWrite("NV.INPUT:");
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else if (oLine.Contains("DLVY"))
                {
                    oWrite(oLine.Replace("DLVY", "SPARE"), true);
                }
                else
                {
                    oWrite("HG" + oLine, true);
                }
            }
            oBlank();
            oWrite("LINK: SCS128_LINE");
            oWrite("ADJUSTABLE ENABLE: 0");
            oWrite("PROTOCOL: SCS.SLAVE");
            oWrite("ADJUSTABLE PORT: 3;");
            oWrite("ADJUSTABLE BAUD: 1200;");
            oWrite("ADJUSTABLE STANDBY.PORT: 4;");
            oWrite("ADJUSTABLE ALTERNATE.BAUD: 1200;");
            oWrite("ADJUSTABLE STOPBITS: 1;");
            oWrite("ADJUSTABLE PARITY: EVEN;");
            oWrite("ADJUSTABLE KEY.ON.DELAY: 50;");
            oWrite("ADJUSTABLE KEY.OFF.DELAY: 50;");
            oWrite("ADJUSTABLE STALE.DATA.TIMEOUT: 60:SEC;");
            oWrite("ADJUSTABLE INTERBYTE.TIMEOUT: 0:MSEC;");
            oWrite("ADJUSTABLE INDICATION.ACK: ENABLED;");
            oWrite("ADJUSTABLE POINT.POINT: 1;");
            oBlank();
            oWrite("ADDRESS: 0");
            oWrite("ADJUSTABLE ENABLE: 1");
            oBlank();
            oWrite("NV.OUTPUT:");
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else
                {
                    oWrite("HS" + oLine, true);
                }
            }
            oBlank();
            oWrite("NV.INPUT:");
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else if (oLine.Contains("DLVY"))
                {
                    oWrite(oLine.Replace("DLVY", "SPARE"), true);
                }
                else
                {
                    oWrite("HS" + oLine, true);
                }
            }
            oBlank();
            oWrite("LINK: ATCS_LINE");
            oWrite("ADJUSTABLE ENABLE: 0");
            oWrite("PROTOCOL: ATCS.SLAVE");
            oWrite("ADJUSTABLE PORT: 3;");
            oWrite("ADJUSTABLE BAUD: 9600;");
            oWrite("ADJUSTABLE POLLING.TIMEOUT: 500:MSEC;");
            oWrite("ADJUSTABLE POLLING.INTERVAL: 1000:MSEC;");
            oWrite("ADJUSTABLE HDLC.FAIL.TIMEOUT: 60:SEC;");
            oWrite("ADJUSTABLE STALE.DATA.TIMEOUT: 120:SEC;");
            oWrite("ADJUSTABLE XMIT.ACK.TIMEOUT:   120:SEC;");
            oWrite("ADJUSTABLE INDICATION.BROADCAST.INTERVAL: 60:SEC;");
            oWrite("ADJUSTABLE MCP.ATCS.ADDRESS: \"78A2AAAAAAA1A1\";");
            oWrite("ADJUSTABLE DEFAULT.ATCS.HOST.ADDRESS: \"28A2AAAAAA\";");
            oWrite("ADJUSTABLE HEALTH.ATCS.ADDRESS: \"AAAAAAAAAA\";");
            oWrite("ADJUSTABLE ADDRESS: \"78A2AAAAAAA2A2\"");
            oBlank();
            oWrite("ADJUSTABLE ENABLE: 1");
            oWrite("STATION.NAME: XOVR;");
            oBlank();
            oWrite("NV.OUTPUT:");
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else
                {
                    oWrite("HAT" + oLine, true);
                }
            }
            oBlank();
            oWrite("NV.INPUT:");
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else if (oLine.Contains("DLVY"))
                {
                    oWrite(oLine.Replace("DLVY", "SPARE"), true);
                }
                else
                {
                    oWrite("HAT" + oLine, true);
                }
            }
            oBlank();
            oWrite("LINK: UCE_LINE");
            oWrite("ADJUSTABLE ENABLE: 0");
            oWrite("PROTOCOL: UCE.SLAVE");
            oWrite("ADJUSTABLE POINT.POINT: 1;");
            oWrite("ADJUSTABLE PORT: 3;");
            oWrite("ADJUSTABLE BAUD: 2400;");
            oWrite("ADJUSTABLE STOPBITS: 1;");
            oWrite("ADJUSTABLE PARITY: NONE;");
            oWrite("ADJUSTABLE ACK.TIMEOUT: 1:SEC;");
            oWrite("ADJUSTABLE XMIT.RETRY.LIMIT: 3;");
            oWrite("ADJUSTABLE STALE.DATA.TIMEOUT: 60:SEC;");
            oWrite("ADJUSTABLE BROADCAST.INTERVAL: 60:SEC;");
            oWrite("ADJUSTABLE BUSY.TIMEOUT: 60:SEC;");
            oBlank();
            oWrite("ADDRESS: 0");
            oWrite("ADJUSTABLE ENABLE: 1");
            oBlank();
            oWrite("NV.OUTPUT:");
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else
                {
                    oWrite("HU" + oLine, true);
                }
            }
            oBlank();
            oWrite("NV.INPUT:");
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else if (oLine.Contains("DLVY"))
                {
                    oWrite(oLine.Replace("DLVY", "SPARE"), true);
                }
                else
                {
                    oWrite("HU" + oLine, true);
                }
            }
            oBlank();
            oWrite("LINK: ML2");
            oWrite("ADJUSTABLE ENABLE: 1");
            oWrite("PROTOCOL: GENISYS.MASTER");
            oWrite("ADJUSTABLE PORT: 1;");
            oWrite("ADJUSTABLE BAUD: 1200;");
            oWrite("ADJUSTABLE STOPBITS: 1;");
            oWrite("ADJUSTABLE PARITY: NONE;");
            oWrite("ADJUSTABLE KEY.ON.DELAY: 12;");
            oWrite("ADJUSTABLE KEY.OFF.DELAY: 12;");
            oWrite("FIXED SECURE.MODE: ON;");
            oWrite("FIXED MASTER.CHECKBACK: ON;");
            oWrite("ADJUSTABLE POINT.POINT: 1;");
            oBlank();
            oWrite("ADDRESS: 1");
            oWrite("ADJUSTABLE ENABLE: 1");
            oBlank();
            oWrite("NV.INPUT:");
            foreach (var oLine in oOutput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else if (oLine.Contains("BKUP"))
                {
                    oWrite(oLine.Replace("BKUP", "SPARE"), true);
                }
                else
                {
                    oWrite("L" + oLine, true);
                }
            }
            oBlank();
            oWrite("NV.OUTPUT:");
            foreach (var oLine in oInput)
            {
                if (oLine.Contains("SPARE"))
                {
                    oWrite(oLine, true);
                }
                else
                {
                    oWrite("L" + oLine, true);
                }
            }
            oBlank();
            oWrite("NV.BOOLEAN BITS");
            oBlank();
            oWrite("DLVY;");
            oBlank();
            oWrite("TIMER BITS");
            oBlank();
            oWrite("FIXED DLVY: SET = 0:SEC CLEAR = 1:SEC;");
            oBlank();
            oWrite("CONFIGURATION");
            oBlank();
            oWrite("SYSTEM");
            oBlank();
            oWrite("ADJUSTABLE DEBUG_PORT_ADDRESS:      1;");
            oWrite("ADJUSTABLE DEBUG_PORT_BAUDRATE:     9600;");
            oWrite("LOGIC_TIMEOUT: 500:MSEC;");
            oBlank();
            oWrite("LOGIC BEGIN");
            oBlank();
            oWrite("NV.ASSIGN DLVY TO LDLVY;");
            oBlank();
            foreach (string oLine in oInput.Where(x => x.Contains("NWZ") || x.Contains("RWZ")))
            {
                oWrite(String.Format("NV.ASSIGN HG{0} + HS{0} + HAT{0} + HU{0} TO L{0};", oLine.Replace(";", "").Replace(",", "")));
            }
            oBlank();
            foreach (string oLine in oInput.Where(x => x.Contains("BLZ")))
            {
                oWrite(String.Format("NV.ASSIGN HG{0} + HS{0} + HAT{0} + HU{0} TO L{0};", oLine.Replace(";", "").Replace(",", "")));
            }
            oBlank();

            GrabSlotOff();
            GrabCOSlotOff();

            //using (SlotForm = new SlotOffSelect())
            //{
            //    SlotForm.TopLevel = true;
            //    SlotForm.TopMost = true;
            //    oWidth = (oInput.OrderByDescending(x => x.Length).First().Length * 6.5);
            //    oVertical = 13;
            //    oHorizontal = 10;
            //    oLabelCount = (from string word in oInput where word.Contains("GZ") select word).Count() * 2;
            //    oSlotOffs = oOutput.Where(x => x.Contains("TK") && !x.Contains("_") && !x.Contains("TEST")).ToList();
            //    oSlotOffs = oSlotOffs.Select(x => NoPunct(x)).ToList();
            //    oLabels = new Label[oLabelCount];
            //    oLabelCount = 0;
            //    foreach (string oLine in oInput.Where(x => x.Contains("GZ")))
            //    {
            //        oLabels[oLabelCount] = new Label();
            //        SlotForm.Controls.Add(oLabels[oLabelCount]);
            //        oLabels[oLabelCount].AutoSize = false;
            //        oLabels[oLabelCount].Height = 13;
            //        oLabels[oLabelCount].Width = (int)oWidth;
            //        oLabels[oLabelCount].Text = NoPunct(oLine) + ":";
            //        oLabels[oLabelCount].TextAlign = ContentAlignment.TopRight;
            //        oLabels[oLabelCount].Location = new System.Drawing.Point(oHorizontal, oVertical);
            //        oLabels[oLabelCount].Name = "GZ" + (oLabelCount + 1);
            //        oLabelCount++;
            //        oLabels[oLabelCount] = new Label();
            //        SlotForm.Controls.Add(oLabels[oLabelCount]);
            //        oLabels[oLabelCount].AutoSize = true;
            //        oCurrent = oSlotOffs.FirstOrDefault(s => s.StartsWith(oLine.Substring(0, 5)));
            //        oCurrent = oSlotOffs.FirstOrDefault(s => s.StartsWith(oLine.Substring(0, 1)));
            //        oCurrent = oSlotOffs.FirstOrDefault(s => s.StartsWith(oLine.Substring(0, 3)));
            //        if (oCurrent == null)
            //        {
            //            oCurrent = oSlotOffs[0];
            //        }
            //        oLabels[oLabelCount].Text = NoPunct(oCurrent);
            //        oHorizontal = oHorizontal + (int)oWidth;
            //        oLabels[oLabelCount].Location = new System.Drawing.Point(oHorizontal, oVertical);
            //        oLabels[oLabelCount].Name = "SLOT" + oLabelCount;
            //        oLabels[oLabelCount].BackColor = SystemColors.ControlLight;
            //        oLabels[oLabelCount].BorderStyle = BorderStyle.Fixed3D;
            //        oLabels[oLabelCount].Click += (sender, EventArgs) => { Label_Click(sender, EventArgs); };
            //        //button.Click += (sender, EventArgs) => { buttonNext_Click(sender, EventArgs, item.NextTabIndex); };
            //        oHorizontal = oHorizontal + 40;
            //        if (oHorizontal > 400)
            //        {
            //            oHorizontal = 10;
            //            oVertical = oVertical + 20;
            //        }
            //        oLabelCount++;
            //    }
            //    oVertical = oVertical + 20;
            //    Button oDone = new Button();
            //    oDone.Click += new EventHandler(DoneButton_Click);
            //    oDone.Name = "DoneButton";
            //    SlotForm.Controls.Add(oDone);
            //    oDone.Text = "Done";
            //    oHorizontal = SlotForm.Width / 2;
            //    oDone.Location = new System.Drawing.Point(oHorizontal - (oDone.Width / 2), oVertical);
            //    SlotForm.ShowDialog();
            //}
            foreach (SigReq sig in oSigReqs)
            {
                oWrite(String.Format("NV.ASSIGN GENISYS_LINE.ENABLED * HG{0} * ~L{1} TO HG{0};", sig.GZ, sig.SlotOff));
                oWrite(String.Format("NV.ASSIGN SCS128_LINE.ENABLED * HS{0} * ~L{1} TO HS{0};", sig.GZ, sig.SlotOff));
                oWrite(String.Format("NV.ASSIGN ATCS_LINE.ENABLED * HAT{0} * ~L{1} TO HAT{0};", sig.GZ, sig.SlotOff));
                oWrite(String.Format("NV.ASSIGN UCE_LINE.ENABLED * HU{0} * ~L{1} TO HU{0};", sig.GZ, sig.SlotOff));
                oBlank();
            }
            foreach (SigReq sig in oSigReqs)
            {
                oWrite(String.Format("NV.ASSIGN HG{0} + HS{0} + HAT{0} + HU{0} TO L{0};", sig.GZ));
            }
            oBlank();
            foreach (SigReq sig in oCOSigReqs)
            {
                oWrite(String.Format("NV.ASSIGN HG{0} + HS{0} + HAT{0} + HU{0} TO L{0};", sig.GZ));
            }
            oBlank();
            if (ListIndex(oInput,"MCZ") > -1)
            {
                oWrite(String.Format("NV.ASSIGN HG{0} + HS{0} + HAT{0} + HU{0} TO L{0};",NoPunct(oInput[ListIndex(oInput, "MCZ")])));
            }
            oBlank();
            if (ListIndex(oInput, "TEST") > -1)
            {
                oWrite(String.Format("NV.ASSIGN HG{0} + HS{0} + HAT{0} + HU{0} TO L{0};", NoPunct(oInput[ListIndex(oInput, "TEST")])));
            }
            oBlank();
            if (ListIndex(oInput, "SNOZ") > -1)
            {
                oWrite(String.Format("NV.ASSIGN HG{0} + HS{0} + HAT{0} + HU{0} TO L{0};", NoPunct(oInput[ListIndex(oInput, "SNOZ")])));
            }
            oBlank();
            foreach (string oLine in oOutput.Where(x => x.Contains("NWK") || x.Contains("RWK")))
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oLine)));
            }
            oBlank();
            foreach (string oLine in oOutput.Where(x => x.Contains("BLK")))
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oLine)));
            }
            oBlank();
            foreach (string oLine in oOutput.Where(x => x.Contains("GK")))
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oLine)));
            }
            oBlank();
            foreach (string oLine in oOutput.Where(x => x.Contains("ASK")))
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oLine)));
            }
            oBlank();
            foreach (string oLine in oOutput.Where(x => x.Contains("TK") && !x.Contains("_") && !x.Contains("TEST")))
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oLine)));
            }
            oBlank();
            foreach (string oLine in oOutput.Where(x => x.Contains("_TK")))
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oLine)));
            }
            oBlank();
            foreach (string oLine in oOutput.Where(x => x.Contains("_AK")))
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oLine)));
            }
            oBlank();
            if (ListIndex(oOutput, "FIBER") > -1)
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oOutput[ListIndex(oOutput, "FIBER")])));
            }
            oBlank();
            foreach (string oLine in oOutput.Where(x => x.Contains("LOPK")))
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oLine)));
            }
            foreach (string oLine in oOutput.Where(x => x.Contains("LOK")))
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oLine)));
            }
            if (ListIndex(oOutput, "HEALTH") > -1)
            {
                oWrite(String.Format("NV.ASSIGN L{0} to HG{0},HS{0},HAT{0},HU{0};", NoPunct(oOutput[ListIndex(oOutput, "HEALTH")])));
            }
            if (ListIndex(oOutput, "K1IK") > -1)
            {
                oWrite(String.Format("NV.ASSIGN L{0} to HG{0},HS{0},HAT{0},HU{0};", NoPunct(oOutput[ListIndex(oOutput, "K1IK")])));
            }
            if (ListIndex(oOutput, "SWIK") > -1)
            {
                oWrite(String.Format("NV.ASSIGN L{0} to HG{0},HS{0},HAT{0},HU{0};", NoPunct(oOutput[ListIndex(oOutput, "SWIK")])));
            }
            if (ListIndex(oOutput, "POK") > -1)
            {
                oWrite(String.Format("NV.ASSIGN L{0} to HG{0},HS{0},HAT{0},HU{0};", NoPunct(oOutput[ListIndex(oOutput, "POK")])));
            }
            if (ListIndex(oOutput, "SNOK") > -1)
            {
                oWrite(String.Format("NV.ASSIGN L{0} to HG{0},HS{0},HAT{0},HU{0};", NoPunct(oOutput[ListIndex(oOutput, "SNOK")])));
            }
            if (ListIndex(oOutput, "GENK") > -1)
            {
                oWrite(String.Format("NV.ASSIGN L{0} to HG{0},HS{0},HAT{0},HU{0};", NoPunct(oOutput[ListIndex(oOutput, "GENK")])));
            }
            if (ListIndex(oOutput, "TESTK") > -1)
            {
                oWrite(String.Format("NV.ASSIGN L{0} to HG{0},HS{0},HAT{0},HU{0};", NoPunct(oOutput[ListIndex(oOutput, "TESTK")])));
            }
            oBlank();
            foreach (string oLine in oOutput.Where(x => x.Contains("_REQK")))
            {
                oWrite(String.Format("NV.ASSIGN L{0} TO HG{0},HS{0},HAT{0},HU{0};", NoPunct(oLine)));
            }
            oBlank();
            oWrite("NV.ASSIGN GENISYS_LINE.STANDBY + SCS128_LINE.STANDBY + ATCS_LINE.STANDBY + UCE_LINE.STANDBY TO LED.8;");
            oBlank();
            oWrite("NV.ASSIGN GENISYS_LINE.STANDBY + SCS128_LINE.STANDBY + ATCS_LINE.STANDBY + UCE_LINE.STANDBY TO HGBKUP, HSBKUP, HATBKUP, HUBKUP;");
            oBlank();
            oWrite("END LOGIC");
            oBlank();
            oWrite("NUMERIC BEGIN");
            oBlank();
            oWrite("BLOCK 1 TRIGGERS ON");
            oWrite("ATCS_LINE.XOVR.INPUTS.RECEIVED,GENISYS_LINE.0.INPUTS.RECEIVED,SCS128_LINE.0.INPUTS.RECEIVED,UCE_LINE.0.INPUTS.RECEIVED AND STALE AFTER 0:SEC;");
            oBlank();
            oWrite("NV.ASSIGN ~DLVY TO DLVY;");
            oWrite("NV.ASSIGN ~DLVY TO DLVY;");
            oBlank();
            oWrite("END BLOCK");
            oBlank();
            oWrite("END NUMERIC");
            oBlank();
            oWrite("END PROGRAM");
            oNewFile = sProgram[0].Replace("GENISYS_II PROGRAM ", "").Replace(";", "");
            oNewFolder = Directory.GetParent(oSourceFolder) + @"\" + oNewFile;
            oCompleted = string.Join(Environment.NewLine, sProgram.ToArray());
        }
        public void Label_Click(object sender, EventArgs e)
        {
            oCurrentLabel = sender as Label;
            var control = SlotForm.Controls.OfType<Label>().FirstOrDefault(c => c.Name == oCurrentLabel.Name.Replace("SLOT", "GZ"));
            TrackForm = new SlotTracks();
            OSBox = new ComboBox();
            oGZLabel = new Label();
            oGZLabel.AutoSize = true;
            TrackForm.Controls.Add(oGZLabel);
            TrackForm.Controls.Add(OSBox);
            oVertical = 13;
            oHorizontal = 10;
            oGZLabel.Text = control.Text;
            oGZLabel.Location = new System.Drawing.Point(oHorizontal, oVertical);
            oVertical = 10;
            oHorizontal = (int)(oHorizontal + oGZLabel.Text.Length * 10);
            OSBox.Location = new System.Drawing.Point(oHorizontal, oVertical);
            OSBox.Width = 60;
            OSBox.DataSource = oSlotOffs;
            OSBox.Text = oCurrentLabel.Text;
            OSBox.SelectedIndexChanged += (osender, EventArgs) => { Chose_Track(osender, EventArgs); };
            TrackForm.TopLevel = true;
            TrackForm.TopMost = true;
            TrackForm.ShowDialog();
        }
        public int ListIndex(List<string> sList, string sString)
        {
            return (sList.FindIndex(x => x.Contains(sString)));
        }
        public void Chose_Track(object sender, EventArgs e)
        {
            oCurrentLabel.Text = OSBox.Text;
            TrackForm.Dispose();
        }
        public void DoneButton_Click(object sender, EventArgs e)
        {
            foreach (Control x in SlotForm.Controls)
            {
                if (x is Label)
                {
                    oGZs.Add(x.Text);
                }
            }
            SlotForm.Dispose();
        }
        public string NoPunct(string sString)
        {
            return sString.Replace(",", "").Replace(";", "").Replace(":","");
        }

        public void oWrite(string sString, bool oIndent = false)
        {
            if (oIndent == true)
            {
                sProgram.Add("  " + sString);
            }
            else
            {
                sProgram.Add(sString);
            }
        }

        public void oBlank()
        {
            sProgram.Add("");
        }

        public void NewDoc(string oRun = "")
        {
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Add();
            oDoc.Content.Text = oCompleted;
            oDoc.Content.set_Style("No Spacing");
            oDoc.Content.Font.Size = 10;
            oDoc.Content.Font.Name = "courier new";
            Directory.CreateDirectory(oNewFolder);
            oDoc.SaveAs2(oNewFolder + @"\" + oNewFile + ".docx");
            oWord.Visible = true;
            if (oRun == "Vital")
            {
                oWord.Run("Exterior.ExteriorRun");
            }
            else if(oRun == "NonVital")
            {
                oWord.Run("Exterior.ExteriorRunNV");
            }
        }
        public void WriteExt()
        {
            NewDoc();
            oWord.Visible = true;
            //oSourceFolder = Directory.GetParent(oSourceFolder) + oProgramName;
            //Directory.CreateDirectory(oSourceFolder);
            //oDoc.SaveAs2(oSourceFolder + @"\" + oProgramName + ".docx");
            //oWord.Run("Exterior.ExteriorRun");
            //oCloseDoc();
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
            oStart = oProgram.FindIndex(x => x.ToUpper().Contains("LOG BITS"));
            if(oStart == -1)
            {
                oStart = oProgram.FindIndex(x => x.ToUpper().Contains("LOG"));
                oStart = oProgram.FindIndex(oStart, x => x.ToUpper().Contains("BITS"));
            }
            oEnd = oProgram.FindIndex(oStart, x => x.ToUpper().Contains(";"));

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
            oNewFile = Path.GetFileNameWithoutExtension(oSourceFile);
            oNewFolder = oSourceFolder;
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
            sProgram.Clear();
            RemoveNotes();
            oProgram = oProgram.Select(x => x.Trim()).ToList();
            oProgram = oProgram.Where(x => x.Trim() != "").ToList();
        }
        public bool ProgramSelect(string oFileType = ".doc")
        {
            OpenFileDialog oProgramSelect = new OpenFileDialog();
            oProgramSelect.Filter = "Microlok Files|*" + oFileType + "*";
            //oProgramSelect.Filter = "Microlok Files|*.doc*";
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
        //public void DocToPlainCODE7248()
        //{
        //    TextExtractor extractor = new TextExtractor(oSourceFile.ToLower());
        //    DocText = extractor.ExtractText();
        //    switch (Path.GetExtension(oSourceFile).ToUpper())
        //    {
        //        case ".DOCX":
        //            stringSeparators = new string[] { "\r\n" };
        //            oProgram = DocText.Split(stringSeparators, StringSplitOptions.None).ToList();
        //            oProgram = oProgram.Select(s => s.Replace("\t", " ")).ToList();
        //            break;
        //        case ".DOC":
        //            stringSeparators = new string[] { "\r" };
        //            oProgram = DocText.Split(stringSeparators, StringSplitOptions.None).ToList();
        //            break;
        //        default:
        //            break;
        //    }
        //    CleanUp();
        //}
        public void oShow()
        {
            Show();
            CenterToScreen();
            this.TopLevel = true;
            this.TopMost = true;
        }
        public void ToolForm_Load(object sender, EventArgs e)
        {

        }
        public void LogBitsButton_Click(object sender, EventArgs e)
        {
            Hide();
            if (ProgramSelect() == true)
            {
                DocToPlain();
                LogBitsBuilder();
                WriteLog();
            }
            oShow();
        }
        public void ExtensionButton_Click(object sender, EventArgs e)
        {
            Hide();
            if (ProgramSelect() == true)
            {
                DocToPlain();
                ExtBuilder();
                WriteExt();
            }
            oShow();
        }
        public void NonVitalButton_Click(object sender, EventArgs e)
        {
            Hide();
            if (ProgramSelect() == true)
            {
                DocToPlain();
                NonVitalBuilder();
                NewDoc("NonVital");
                oCloseDoc();
            }
            oShow();
        }
        public void MLLConvertButton_Click(object sender, EventArgs e)
        {
            Hide();
            if (ProgramSelect(".mll") == true)
            {
                File.Copy(oSourceFile, oSourceFolder + @"\" + Path.GetFileNameWithoutExtension(oSourceFile) + "-Backup.mll",true);
                MllStripper();
                NewDoc("Vital");
                oCloseDoc();
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
        private void QLCPButton_Click(object sender, EventArgs e)
        {
            Hide();
            if (ProgramSelect() == true)
            {
                DocToPlain();
                QLCPBuilder();
                NewDoc("Vital");
                oCloseDoc();
            }
            oShow();
        }
        private void NotesButton_Click(object sender, EventArgs e)
        {
            Hide();
            if (ProgramSelect() == true)
            {
                DocToPlain();
                oCompleted = string.Join(Environment.NewLine, oProgram.ToArray());
                oNewFile = Path.GetFileNameWithoutExtension(oSourceFile) + "_NoNotes";
                oNewFolder = Directory.GetParent(oSourceFolder) + @"\" + oNewFile;
                NewDoc();
            }
            oShow();
        }
        private void SwitchButton_Click(object sender, EventArgs e)
        {
            Hide();
            if (ProgramSelect() == true)
            {
                DocToPlain();
                CheckSwitches();
            }
            oShow();
        }
        private void BooleanButton_Click(object sender, EventArgs e)
        {
            Hide();
            if (ProgramSelect() == true)
            {
                DocToPlain();
                CreateBoolean();
            }
            oShow();
        }
    }
}