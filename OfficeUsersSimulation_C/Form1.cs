/**************************************************************************************
    OfficeUsersSimulation
    Copyright (C) 2017  Apostolos Smyrnakis - IT/CDA/AD - apostolos.smyrnakis@cern.ch

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 **************************************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using WindowsApplication1;
using OfficeUsersSimulation_C;

namespace OfficeUsersSimulation_C
{
    public partial class Form1 : Form
    {
        // ------------------------------------ Declarations ------------------------------------
        public string workingDirectory;
        public bool workingPathOK = false;
        public int filesToCreate = 0;
        public bool usingWordFiles = false;
        public bool usingExcelFiles = false;
        public bool addDelay = false;
        public int addDelaySec = 0;
        public bool autoDelete = false;
        public bool varyFormatting = false;

        public bool programRunning = false;
        public bool emptyFilesRunning = false;
        bool intialLoaded = false;

        public List<string> listOfWordFiles = new List<string>();
        public List<string> listOfExcelFiles = new List<string>();
        public List<string> listOfEmptyWordFiles = new List<string>();
        public List<string> listOfEmptyExcelFiles = new List<string>();

        public string wordFileName = "testWord_";
        public string excelFileName = "testExcel_";
        public string emptyFileName = "emptyFile_";

        public string textToSearch1 = "sit";
        public string textToSearch2 = "magna";
        public string textToReplace1 = "Apostolos";
        public string textToReplace2 = "Smyrnakis";
        
        Stopwatch timmerWordFiles = new Stopwatch();
        Stopwatch timmerExcelFiles = new Stopwatch();
        Stopwatch timmerEmptyFiles = new Stopwatch();

        public string[] lorems = {loremsClass.lorem1, loremsClass.lorem5, loremsClass.lorem10, loremsClass.lorem15, loremsClass.lorem20 };
        // ---------------------------------------------------------------------------------------

        // ------------ Initialize values & items - Restore last settings from memory ------------
        public Form1()
        {
            var winwordCheck = new Word.Application();
            if (winwordCheck == null)
                MessageBox.Show("Microsoft World could NOT be located on your system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            winwordCheck.Quit();
            if (winwordCheck != null) Marshal.ReleaseComObject(winwordCheck);
            winwordCheck = null;
            var winexcelCheck = new Microsoft.Office.Interop.Excel.Application();
            if (winexcelCheck == null)
                MessageBox.Show("Microsoft Excel could NOT be located on your system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            winexcelCheck.Quit();
            if (winexcelCheck != null) Marshal.ReleaseComObject(winexcelCheck);
            winexcelCheck = null;
            InitializeComponent();
            textBoxWorkingDir.ForeColor = System.Drawing.Color.Red;
            textBoxRuntime1.Visible = textBoxRuntime2.Visible = false;                  // Hide textBox1,5 (used for logging durring program running)
            this.FormBorderStyle = FormBorderStyle.Fixed3D;
            progressBar1.MarqueeAnimationSpeed = 10;                                    // Set moving speed for progressBar
            progressBar1.Style = ProgressBarStyle.Continuous;                           // Disable progressBar movement
            //loadSavedSettings();
            //checkLoadedPaths();
            ActiveControl = trackBar1;          // This should be always the last one, after setting all other properties!
        }
        // ---------------------------------------------------------------------------------------

        // --------------------------- Working Directory - Sample Files --------------------------
        private void loadSavedSettings()
        {
            //MessageBox.Show("Pre-saved settings are loaded!", "loadSavedSettings()", MessageBoxButtons.OK, MessageBoxIcon.Information); // for debug!!!
            textBoxWorkingDir.Text = Properties.Settings.Default["lastWorkingDir"].ToString();   // Load last working directory
            workingDirectory = textBoxWorkingDir.Text;
            if (Properties.Settings.Default.lastSliderValue >= 4 && Properties.Settings.Default.lastSliderValue <= 1000)
                trackBar1.Value = Properties.Settings.Default.lastSliderValue;              // load last number of created files
            else
                trackBar1.Value = 4;
            numericUpDownFilesToCreate.Value = trackBar1.Value;
            filesToCreate = Properties.Settings.Default.lastSliderValue;                // load type of file to be created
            usingWordFiles = Properties.Settings.Default.useWords;                      //      (word and/or excel)
            checkBoxUsingWord.Checked = usingWordFiles ? true : false;
            usingExcelFiles = Properties.Settings.Default.useExcels;
            checkBoxUsingExcel.Checked = usingExcelFiles ? true : false;
            varyFormatting = Properties.Settings.Default.lastVaryFormatting;            // load last vary-formating option
            checkBoxVaryFormating.Checked = varyFormatting ? true : false;
            autoDelete = Properties.Settings.Default.lastAutoDelete;                    // load last auto-delete option
            checkBoxAutoDelete.Checked = autoDelete ? true : false;
            numericUpDownDelayInEditing.Value = Properties.Settings.Default.lastDelayInEditing;      // load last Delay-In-Editing value
            addDelaySec = Convert.ToInt32(numericUpDownDelayInEditing.Value);
            if (addDelaySec == 0)
            {
                checkBoxDelayInEditing.Checked = false;
                addDelay = false;
            }
            else
            {
                checkBoxDelayInEditing.Checked = true;
                addDelay = true;
            }
        }
        // ---------------------------------------------------------------------------------------

        // ----------------------------------- Working Directory --------------------------------- 
        // workingDirectory
        private void button2_Click(object sender, EventArgs e)
        {
            textBoxWorkingDir.ForeColor = System.Drawing.Color.Red;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBoxWorkingDir.Text = folderBrowserDialog1.SelectedPath + "\\";
                    workingDirectory = folderBrowserDialog1.SelectedPath + "\\";
                    Properties.Settings.Default["lastWorkingDir"] = folderBrowserDialog1.SelectedPath + "\\";
                    Properties.Settings.Default.Save();
                    checkLoadedPaths();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error loading Working directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBoxWorkingDir.ForeColor = System.Drawing.Color.Red;
        }
        private void textBox2_Leave(object sender, EventArgs e)
        {
            textBoxWorkingDir.ForeColor = System.Drawing.Color.Red;
            while (textBoxWorkingDir.Text.StartsWith(" "))
                textBoxWorkingDir.Text = textBoxWorkingDir.Text.Substring(1);
            while (textBoxWorkingDir.Text.EndsWith(" "))
                textBoxWorkingDir.Text = textBoxWorkingDir.Text.Substring(0, textBoxWorkingDir.Text.Length - 1);
            if (!textBoxWorkingDir.Text.EndsWith("\\"))
                textBoxWorkingDir.Text += "\\";
            workingDirectory = textBoxWorkingDir.Text;
            Properties.Settings.Default["lastWorkingDir"] = textBoxWorkingDir.Text;
            Properties.Settings.Default.Save();
            checkLoadedPaths();
        }
        // ---------------------------------------------------------------------------------------

        // ------------------ Check if working directory & sample files do exist ----------------- 
        public int checkLoadedPaths()
        {
            if (Directory.Exists(workingDirectory))
            {
                createEmpyFilesToolStripMenuItem.Enabled = true;
                textBoxWorkingDir.ForeColor = System.Drawing.Color.ForestGreen;
                workingPathOK = true;
                return 0;
            }
            else
            {
                createEmpyFilesToolStripMenuItem.Enabled = false;
                textBoxWorkingDir.ForeColor = System.Drawing.Color.Red;
                workingPathOK = false;
                return -1;
            }
        }
        // ---------------------------------------------------------------------------------------

        // ---------------------------- Auto delete created files --------------------------------
        public void autoDeleteHandler()
        {
            if (usingWordFiles)
            {
                try
                {
                    string deletingDir = workingDirectory + "officeSimulation_word\\";
                    System.IO.Directory.Delete(deletingDir, true);
                    listOfWordFiles.Clear();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error deleting \\officeSimulation_word\\");
                    MessageBox.Show(ex.Message, "Error deleting \\officeSimulation_word\\", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (usingExcelFiles)
            {
                try
                {
                    string deletingDir = workingDirectory + "officeSimulation_excel\\";
                    System.IO.Directory.Delete(deletingDir, true);
                    listOfExcelFiles.Clear();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error deleting \\officeSimulation_excel\\");
                    MessageBox.Show(ex.Message, "Error deleting \\officeSimulation_excel\\", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        // ---------------------------------------------------------------------------------------

        // ------------------------ Number of files to create handler ----------------------------
        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            trackBar1.Value = Convert.ToInt32(numericUpDownFilesToCreate.Value);
            filesToCreate = Convert.ToInt32(numericUpDownFilesToCreate.Value);
            Properties.Settings.Default["lastSliderValue"] = trackBar1.Value;
            Properties.Settings.Default.Save();
        }
        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            numericUpDownFilesToCreate.Value = trackBar1.Value;
            filesToCreate = trackBar1.Value;
            Properties.Settings.Default["lastSliderValue"] = trackBar1.Value;
            Properties.Settings.Default.Save();
        }
        // ---------------------------------------------------------------------------------------

        // ---------------------------- Mouse Hover event for Form1 ------------------------------ 
        private void Form1_MouseHover(object sender, EventArgs e)
        {
            if (!intialLoaded)
            {
                loadSavedSettings();
                checkLoadedPaths();
                intialLoaded = true;
            }
        }
        // ---------------------------------------------------------------------------------------

        // ------------------------ AutoDelete & restore default button -------------------------- 
        public void restoreAfterRun()
        {
            if (autoDelete)
            {
                textBoxRuntime2.Text = "Deleting created files...";
                autoDeleteHandler();
                textBoxRuntime2.Text = "Deleting created files ---> DONE!";
            }
            textBoxRuntime1.Clear();
            textBoxRuntime2.Clear();
            textBoxRuntime1.Visible = textBoxRuntime2.Visible = false;
            numericUpDownFilesToCreate.Enabled = true;
            trackBar1.Enabled = true;
            buttonRun.Text = "R U N";
            buttonRun.ForeColor = System.Drawing.Color.ForestGreen;
            programRunning = false;
            if (!backgroundWorkerEmptyFiles.IsBusy)
                progressBar1.Style = ProgressBarStyle.Continuous;
        }
        // ---------------------------------------------------------------------------------------
        
        // ------------------------ Message Box with all variable values ------------------------- 
        private void debuggingMsgBox()
        {
            String msgBoxAboutCaption = "Debugging messageBox";
            String msgBoxAboutText = "filesToCreate: " + filesToCreate + "\r\n\r\nworkingDirectory: " + workingDirectory + "\r\n\r\nusingWordFiles: " + usingWordFiles + "\r\nusingExcelFiles: " + usingExcelFiles + "\r\n\r\naddDelay: " + addDelay + "\r\naddDelaySec: " + addDelaySec + "\r\n\r\nworkingPathOK: " + workingPathOK + "\r\n\r\nautoDelete: " + autoDelete + "\r\nvaryFormatting: " + varyFormatting;
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons);
        }
        // ---------------------------------------------------------------------------------------

        // -------------------------- TextBox1 (up) text during runtime --------------------------
        private void textBox1TextBuild()
        {
            string typeOfFiles = "";
            if (usingWordFiles && usingExcelFiles)
                typeOfFiles = " Word & Excel";
            else if (usingWordFiles)
                typeOfFiles = " Word";
            else if (usingExcelFiles)
                typeOfFiles = " Excel";

            string formatting = "";
            formatting = varyFormatting ? "various" : "no";

            string delay = "";
            if (addDelaySec == 0)
                delay = "No";
            else
                delay = addDelaySec.ToString() + " seconds";

            string deleteAfter = "";
            deleteAfter = autoDelete ? "" : "NOT";

            textBoxRuntime1.Text = "Creating " + filesToCreate + " " + typeOfFiles + " files, with " + formatting + " formatting." + "\r\n" + delay + " seconds delay during each file creation." + "\r\nFiles will " + deleteAfter + " be deleted after operations.";
        }
        // ---------------------------------------------------------------------------------------

        // ------------------------- TextBox5 (down) text during runtime ------------------------- // <?><?><?><?><?><?><?><?><?><?><?><?><?><?><?><?>
        private void textBox5TextBuild()
        {
            textBoxRuntime2.Text = "textBox5 : Under construction! \r\nContact Apostolos for more info.";
        }
        // ---------------------------------------------------------------------------------------

        // --- Button RUN click ------------------------------------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            int pathsCheckResult = checkLoadedPaths(); // using pathsOk variable below
            if (!checkBoxUsingWord.Checked && !checkBoxUsingExcel.Checked)
            {
                MessageBox.Show("Please select at least one between 'Word' or 'Excel' files!");
            }
            else if (!workingPathOK)
            {
                MessageBox.Show("Please check working directory!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (programRunning)        // canceling all background workers & restoring defaults
            {
                cancelBackgroundWorkers();
                Thread.Sleep(200);
                restoreAfterRun();
                MessageBox.Show("Program terminated by the user!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                restoreAfterRun();
            }
            else if (!programRunning)
            {
                programRunning = true;
                textBoxRuntime1.Visible = textBoxRuntime2.Visible = true;
                textBoxRuntime1.BringToFront();
                textBoxRuntime2.BringToFront();
                numericUpDownFilesToCreate.Enabled = false;
                trackBar1.Enabled = false;
                textBox1TextBuild();
                //textBox5TextBuild();
                buttonRun.ForeColor = System.Drawing.Color.Red;
                buttonRun.Text = "S T O P";
                //debuggingMsgBox();          // for debug!!!
                try
                {
                    progressBar1.Style = ProgressBarStyle.Marquee;
                    if (checkBoxUsingWord.Checked)
                    {
                        timmerWordFiles.Reset();
                        timmerWordFiles.Start();
                        listOfWordFiles.Clear();                    // needed if I click run for a 2nd time before closing the GUI
                        backgroundWorkerWordCreate.RunWorkerAsync();
                    }
                    if (checkBoxUsingExcel.Checked)
                    {
                        timmerExcelFiles.Reset();
                        timmerExcelFiles.Start();
                        listOfExcelFiles.Clear();                   // needed if I click run for a 2nd time before closing the GUI
                        backgroundWorkerExcelCreate.RunWorkerAsync();
                    }
                }
                catch (Exception ex)
                {
                    cancelBackgroundWorkers();
                    restoreAfterRun();
                    progressBar1.Style = ProgressBarStyle.Continuous;
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                cancelBackgroundWorkers();
                restoreAfterRun();
                progressBar1.Style = ProgressBarStyle.Continuous;
                MessageBox.Show("Error! 'else' case executed in button1_Click! \r\n Please contact Apostolos Smyrnakis!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // ---------------------------------------------------------------------------------------

        // --- Create empty files menu click -----------------------------------------------------
        private void createEmpyFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (emptyFilesRunning)
            {
                backgroundWorkerEmptyFiles.CancelAsync();
            }
            else
            {
                if (!checkBoxUsingWord.Checked && !checkBoxUsingExcel.Checked)
                {
                    MessageBox.Show("Please select at least one between 'Word' or 'Excel' files!");
                }
                else
                {
                    int pathCheck = checkLoadedPaths();
                    if (pathCheck == 0) // Working directory exists
                    {
                        emptyFilesRunning = true;
                        createEmpyFilesToolStripMenuItem.BackColor = Color.FromKnownColor(KnownColor.ActiveCaption);
                        timmerEmptyFiles.Reset();
                        timmerEmptyFiles.Start();
                        //listOfEmptyWordFiles.Clear();
                        //listOfEmptyExcelFiles.Clear();
                        backgroundWorkerEmptyFiles.RunWorkerAsync();
                        progressBar1.Style = ProgressBarStyle.Marquee;
                        //MessageBox.Show("Creating " + filesToCreate + " empty files!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Working directory not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        // ---------------------------------------------------------------------------------------
        
        // -------------------------------- Copy Excel method ------------------------------------ // <?><?><?><?><?><?><?><?><?><?><?><?><?><?><?><?>
        /*public void CopyExcel(string fromExcelPath, string toExcelPath)
        {
            timmerCopyExcels.Reset();
            timmerCopyExcels.Start();

            //Create an instance for excel app
            Microsoft.Office.Interop.Excel.Application excelCopy = null;
            Microsoft.Office.Interop.Excel.Workbook workBookCopy;
            Microsoft.Office.Interop.Excel.Worksheet workSheetCopy;
            excelCopy.Visible = false;
            excelCopy.DisplayAlerts = false;

            try
            {

                workBookCopy = excelCopy.Workbooks.Open(fromExcelPath);
                workSheetCopy = (excelCopy.Worksheet)WorkBookCopy.WorkSheets.get_Item((int)table + 1);

                var originalDocument = excelCopy.ThisWorkbook.ActiveSheet.  Documents.Open(fromExcelPath);    // Open original document

                originalDocument.ActiveWindow.Selection.WholeStory();               // Select all in original document
                var originalText = originalDocument.ActiveWindow.Selection;         // Copy everything to the variable

                var newDocument = new Word.Document();                              // Create new Word document
                newDocument.Range().Text = originalText.Text;                       // Pasete everything from the variable
                newDocument.SaveAs(toExcelPath); // maybe SaveAs2??                  // Save the new document

                originalDocument.Close(false);
                newDocument.Close();

                excelCopy.Quit();
                excelCopy = null;
            }
            catch (Exception ex)
            {
                timmerCopyExcels.Stop();
                timmerCopyExcels.Reset();
                Console.WriteLine("Error in CopyExcel");
                MessageBox.Show(ex.Message, "Error in CopyExcel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }*/
        // ---------------------------------------------------------------------------------------
        
        // ----------------------------- Count words in a string --------------------------------- 
        private int wordCounter(string textToCount)
        {
            var text = textToCount.Trim();
            int wordCount = 0;
            int index = 0;

            while (index < text.Length)
            {
                // check if current char is part of a word
                while (index < text.Length && !char.IsWhiteSpace(text[index]))
                    index++;

                wordCount++;

                // skip whitespace until next word
                while (index < text.Length && char.IsWhiteSpace(text[index]))
                    index++;
            }
            return wordCount;
        }
        // ---------------------------------------------------------------------------------------

        // --------------------------------- Check Boxes settings -------------------------------- 
        // checkBox: use Word files 
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            usingWordFiles = checkBoxUsingWord.Checked ? true : false;
            Properties.Settings.Default["useWords"] = checkBoxUsingWord.Checked;
            Properties.Settings.Default.Save();
        }

        // checkBox: use Excel files
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            usingExcelFiles = checkBoxUsingExcel.Checked ? true : false;
            Properties.Settings.Default["useExcels"] = checkBoxUsingExcel.Checked;
            Properties.Settings.Default.Save();
        }

        // checkBox: add some seconds delay during document edit
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            addDelay = checkBoxDelayInEditing.Checked ? true : false;
            addDelaySec = Convert.ToInt32(numericUpDownDelayInEditing.Value);
            if (!addDelay)
                numericUpDownDelayInEditing.Value = 0;
            Properties.Settings.Default["lastDelayInEditing"] = Convert.ToInt32(numericUpDownDelayInEditing.Value);
            Properties.Settings.Default.Save();
        }

        // checkBox: auto delete created files
        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            autoDelete = checkBoxAutoDelete.Checked ? true : false;
            Properties.Settings.Default["lastAutoDelete"] = checkBoxAutoDelete.Checked;
            Properties.Settings.Default.Save();
        }

        // checkBox: vary documents formatting
        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            varyFormatting = checkBoxVaryFormating.Checked ? true : false;
            Properties.Settings.Default["lastVaryFormatting"] = checkBoxVaryFormating.Checked;
            Properties.Settings.Default.Save();
        }
        // ---------------------------------------------------------------------------------------

        // ---------------------------------- Other items check ---------------------------------- 
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            checkBoxDelayInEditing.Checked = true;
            addDelaySec = Convert.ToInt32(numericUpDownDelayInEditing.Value);
            if (numericUpDownDelayInEditing.Value == 0)
                checkBoxDelayInEditing.Checked = addDelay = false;
        }
        // ---------------------------------------------------------------------------------------

        // ------------------------------------- Background workers --------------------------------------
        // bgw WORDCreate doWork
        private void backgroundWorkerWordCreate_DoWork(object sender, DoWorkEventArgs e)
        {
            if (!backgroundWorkerWordCreate.CancellationPending)
            {
                int filesToCreateTemp = filesToCreate;             // avoid change in '#files to be created' during runtime
                try
                {
                    string workingDir = workingDirectory + "officeSimulation_word\\";
                    System.IO.Directory.CreateDirectory(workingDir);
                    int documentCount = 0;
                    string wordFileNameDate = wordFileName + DateTime.Now.ToString("dd-MM-yyyy_HH-mm");
                    string wordFileNameDateFinal;

                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~ Word instance ~~~~~~~~~~~~~~~~~~~~~~~~~
                    var winword = new Microsoft.Office.Interop.Word.Application();
                    winword.ShowAnimation = false;
                    winword.Visible = false;
                    object winWordMissing = System.Reflection.Missing.Value;
                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                    while ((documentCount < filesToCreateTemp) && (!backgroundWorkerWordCreate.CancellationPending))
                    {
                        wordFileNameDateFinal = wordFileNameDate + "_" + documentCount;
                        // ~~~~~~ create word files (vary formatting)
                        try
                        {
                            if (varyFormatting)
                            {
                                //Selecting random lorem text length
                                Random r = new Random(DateTime.Now.Millisecond);
                                int rLorem1 = r.Next(0, 5);
                                int rLorem2 = r.Next(0, 5);
                                int rLorem3 = r.Next(0, 5);
                                int rLorem4 = r.Next(0, 2);
                                int rLorem5 = r.Next(0, 2);

                                //Create a new document
                                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref winWordMissing, ref winWordMissing, ref winWordMissing, ref winWordMissing);

                                //Add header into the document
                                foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                                {
                                    //Get the header range and add the header details.
                                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                                    headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                                    headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                    headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                                    headerRange.Font.Size = 10;
                                    headerRange.Text = lorems[rLorem4];
                                }

                                //Add the footers into the document
                                foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                                {
                                    //Get the footer range and add the footer details.
                                    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                                    footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                                    footerRange.Font.Size = 10;
                                    footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                    footerRange.Text = lorems[rLorem5];
                                }

                                //adding text to document
                                document.Content.SetRange(0, 0);
                                document.Content.Text = lorems[rLorem1] + "\r\n";
                                // adding some delay
                                if (addDelay)
                                    Thread.Sleep(addDelaySec * 1000);
                                document.Content.Text = lorems[rLorem2] + "\r\n";
                                document.Content.Text = lorems[rLorem3];

                                //Add paragraph with Heading 1 style
                                Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref winWordMissing);
                                object styleHeading1 = "Heading 1";
                                para1.Range.set_Style(ref styleHeading1);
                                para1.Range.Text = lorems[rLorem1];
                                para1.Range.InsertParagraphAfter();

                                //Add paragraph with Heading 2 style
                                Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref winWordMissing);
                                object styleHeading2 = "Heading 2";
                                para2.Range.set_Style(ref styleHeading2);
                                para2.Range.Text = lorems[rLorem2];
                                para2.Range.InsertParagraphAfter();

                                //Save the document
                                object filename = workingDir + wordFileNameDateFinal + ".docx";
                                //object filename = filePath + "\\" + fileName + ".docx";
                                document.SaveAs2(ref filename);
                                // Keeping a list of files created
                                listOfWordFiles.Add(filename.ToString());
                                document.Close(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                                document = null;
                                if (document != null) Marshal.ReleaseComObject(document);

                            }
                            else // ~~~~~~ create word files (NO formatting)
                            {
                                //adding text to document
                                Random r = new Random(DateTime.Now.Millisecond);
                                int rLorem1 = r.Next(0, 5);
                                int rLorem2 = r.Next(0, 5);
                                int rLorem3 = r.Next(0, 5);

                                //Create a new document
                                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref winWordMissing, ref winWordMissing, ref winWordMissing, ref winWordMissing);

                                document.Content.SetRange(0, 0);
                                document.Content.Text = lorems[rLorem1] + "\r\n";
                                // adding some delay
                                if (addDelay)
                                    Thread.Sleep(addDelaySec * 1000);
                                document.Content.Text = lorems[rLorem2] + "\r\n";
                                document.Content.Text = lorems[rLorem3];

                                //Save the document
                                object filename = workingDir + wordFileNameDateFinal + ".docx";
                                //object filename = filePath + "\\" + fileName + ".docx";
                                document.SaveAs2(ref filename);
                                // Keeping a list of files created
                                listOfWordFiles.Add(filename.ToString());
                                document.Close(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                                document = null;
                                if (document != null) Marshal.ReleaseComObject(document);
                            }
                        }
                        catch (Exception ex)
                        {
                            backgroundWorkerWordCreate.CancelAsync();
                            MessageBox.Show(ex.Message, "Error creating Word file!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        documentCount++;
                    }

                    // ~~~~~~ copy some of the documents
                    int i2 = 0;
                    while (i2 < filesToCreateTemp && !backgroundWorkerWordCreate.CancellationPending)
                    {
                        if (i2 % 4 == 0)
                        {
                            string filePathNoExtension = listOfWordFiles[i2].Substring(0, listOfWordFiles[i2].Length - 5);
                            string newFilePathCopy = filePathNoExtension + "_copy" + ".docx";
                            try
                            {
                                var originalDocument = winword.Documents.Open(listOfWordFiles[i2]);    // Open original document

                                originalDocument.ActiveWindow.Selection.WholeStory();               // Select all in original document
                                var originalText = originalDocument.ActiveWindow.Selection;         // Copy everything to the variable

                                var newDocument = new Word.Document();                              // Create new Word document
                                newDocument.Range().Text = originalText.Text;                       // Pasete everything from the variable
                                newDocument.SaveAs(newFilePathCopy); // maybe SaveAs2??                  // Save the new document

                                originalDocument.Close(false);
                                newDocument.Close();
                                
                                if (originalDocument != null) Marshal.ReleaseComObject(originalDocument);
                                if (newDocument != null) Marshal.ReleaseComObject(newDocument);
                            }
                            catch (Exception ex)
                            {
                                backgroundWorkerWordCreate.CancelAsync();
                                MessageBox.Show(ex.Message, "Error coping documents", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        i2++;
                    }

                    // ~~~~~~ find-replace in some documents
                    int i3 = 0;
                    while (i3 < filesToCreateTemp && !backgroundWorkerWordCreate.CancellationPending)
                    {
                        if (i3 % 3 == 0)
                        {
                            try
                            {
                                Microsoft.Office.Interop.Word.Document aDoc = winword.Documents.Open(listOfWordFiles[i3], ReadOnly: false, Visible: false);
                                aDoc.Activate();
                                
                                winword.Selection.Find.Execute(textToSearch1, false, true, false, false, false, true, 1, false, textToReplace1, 2, false, false, false, false);
                                winword.Selection.Find.Execute(textToSearch2, false, true, false, false, false, true, 1, false, textToReplace2, 2, false, false, false, false);

                                aDoc.SaveAs2();

                                aDoc.Close(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                                aDoc = null;
                                if (aDoc != null) Marshal.ReleaseComObject(aDoc);
                            }
                            catch (Exception ex)
                            {
                                backgroundWorkerWordCreate.CancelAsync();
                                MessageBox.Show(ex.Message, "Error in find-replace", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        i3++;
                    }

                    // ~~~~~~~~~~~~~~~~~~~~ Terminating Word instance ~~~~~~~~~~~~~~~~~~~
                    winword.Quit(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                    winword.Quit();
                    if (winword != null) Marshal.ReleaseComObject(winword);
                    winword = null;
                    winWordMissing = null;
                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                }
                catch (Exception ex)
                {
                    backgroundWorkerWordCreate.CancelAsync();
                    MessageBox.Show(ex.Message, "backGroundWorker WordCreate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                backgroundWorkerWordCreate.Dispose();
                backgroundWorkerWordCreate.RunWorkerCompleted -= new RunWorkerCompletedEventHandler(backgroundWorkerWordCreate_RunWorkerCompleted);
                backgroundWorkerWordCreate.DoWork -= new DoWorkEventHandler(backgroundWorkerWordCreate_DoWork);
            }
        }
        // bgw WORDCreate Completed
        private void backgroundWorkerWordCreate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (programRunning)
            {
                timmerWordFiles.Stop();
                TimeSpan runTimeWordFiless = timmerWordFiles.Elapsed;
                timmerWordFiles.Reset();
                MessageBox.Show("Word files created!\r\n\r\nRun time: " + runTimeWordFiless.TotalSeconds.ToString().Substring(0, 6) + " seconds.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (!backgroundWorkerExcelCreate.IsBusy)
                restoreAfterRun();
        }

        // bgw EXCELCreate doWork
        private void backgroundWorkerExcelCreate_DoWork(object sender, DoWorkEventArgs e)
        {
            if (!backgroundWorkerExcelCreate.CancellationPending)
            {
                int filesToCreateTemp = filesToCreate;             // avoid change in #files to be created during runtime
                try
                {
                    string workingDir = workingDirectory + "officeSimulation_excel\\";
                    System.IO.Directory.CreateDirectory(workingDir);
                    int excelCount = 0;
                    string excelFileNameDate = excelFileName + DateTime.Now.ToString("dd-MM-yyyy_HH-mm");
                    string excelFileNameDateFinal;

                    // ~~~~~~~~~~~~~~~~~~~~~~~~~ Excel instance ~~~~~~~~~~~~~~~~~~~~~~~~~
                    Microsoft.Office.Interop.Excel.Application excel;
                    Microsoft.Office.Interop.Excel.Workbook workBook;
                    Microsoft.Office.Interop.Excel.Worksheet workSheet;
                    excel = new Microsoft.Office.Interop.Excel.Application();
                    excel.Visible = false;
                    excel.DisplayAlerts = false;
                    object m = Type.Missing;
                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                    while ((excelCount < filesToCreateTemp) && (!backgroundWorkerExcelCreate.CancellationPending))
                    {
                        //Selecting random rows and columns count
                        Random r = new Random(DateTime.Now.Millisecond);
                        int rowsToWrite = r.Next(10, 70);
                        int colsToWrite = r.Next(10, 30);

                        excelFileNameDateFinal = excelFileNameDate + "_" + excelCount;
                        try
                        {
                            // ~~~~~~ create excel files (vary formatting)
                            if (varyFormatting)
                            {
                                //Create a new WorkBook & WorkSheet
                                workBook = excel.Workbooks.Add(Type.Missing);
                                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                                workSheet.Name = "VaryFormatting";

                                var data = new object[rowsToWrite, colsToWrite];
                                for (var row = 1; row <= rowsToWrite; row++)
                                {
                                    if (row == 5)
                                    {
                                        if (addDelay)
                                            Thread.Sleep(addDelaySec * 1000);   // delay in editing
                                    }
                                    for (var column = 1; column <= colsToWrite; column++)
                                    {
                                        data[row - 1, column - 1] = r.Next(99, 99999);
                                    }
                                }

                                var startCell = (Range)workSheet.Cells[1, 1];
                                var endCell = (Range)workSheet.Cells[rowsToWrite, colsToWrite];
                                var writeRange = workSheet.Range[startCell, endCell];

                                writeRange.Value2 = data;
                                writeRange.Cells.Font.Size = 25;
                                workSheet.Range[(Range)workSheet.Cells[1, 1], (Range)workSheet.Cells[rowsToWrite / 2, colsToWrite / 2]].Font.Italic = true;
                                workSheet.Range[(Range)workSheet.Cells[1, 1], (Range)workSheet.Cells[rowsToWrite / 2, colsToWrite / 2]].Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightBlue);
                                workSheet.Range[(Range)workSheet.Cells[rowsToWrite / 2 + 1, colsToWrite / 2 + 1], (Range)workSheet.Cells[rowsToWrite, colsToWrite]].Font.Bold = true;
                                workSheet.Range[(Range)workSheet.Cells[rowsToWrite / 2 + 1, colsToWrite / 2 + 1], (Range)workSheet.Cells[rowsToWrite, colsToWrite]].Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightGreen);
                                workSheet.Range[(Range)workSheet.Cells[1, colsToWrite / 2 + 1], (Range)workSheet.Cells[rowsToWrite / 2, colsToWrite]].Font.Color = System.Drawing.ColorTranslator.ToOle(Color.DarkOrange);
                                workSheet.Range[(Range)workSheet.Cells[rowsToWrite / 2 + 1, 1], (Range)workSheet.Cells[rowsToWrite, colsToWrite / 2]].Font.Color = System.Drawing.ColorTranslator.ToOle(Color.DarkViolet);
                            }
                            // ~~~~~~ create excel files (NO formatting)
                            else
                            {
                                //Create a new WorkBook & WorkSheet
                                workBook = excel.Workbooks.Add(Type.Missing);
                                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                                workSheet.Name = "PlainText";

                                var data = new object[rowsToWrite, colsToWrite];
                                for (var row = 1; row <= rowsToWrite; row++)
                                {
                                    if (row == 5)
                                    {
                                        if (addDelay)
                                            Thread.Sleep(addDelaySec * 1000);
                                    }
                                    for (var column = 1; column <= colsToWrite; column++)
                                    {
                                        data[row - 1, column - 1] = r.Next(99, 99999);
                                    }
                                }

                                var startCell = (Range)workSheet.Cells[1, 1];
                                var endCell = (Range)workSheet.Cells[rowsToWrite, colsToWrite];
                                var writeRange = workSheet.Range[startCell, endCell];

                                writeRange.Value2 = data;
                            }
                            object filename = workingDir + excelFileNameDateFinal + ".xlsx";
                            // Keeping a list of files created
                            workBook.SaveAs(filename);
                            listOfExcelFiles.Add(filename.ToString());
                            workBook.Close();
                            workBook = null;
                            workSheet = null;
                            if (workSheet != null) Marshal.ReleaseComObject(workSheet);
                            if (workBook != null) Marshal.ReleaseComObject(workBook);
                        }
                        catch (Exception ex)
                        {
                            backgroundWorkerExcelCreate.CancelAsync();
                            MessageBox.Show(ex.Message, "Error creating Excel file!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        excelCount++;
                    }

                    // ~~~~~~ copy some of the excels
                    /*int i2 = 0;
                    while (i2 < filesToCreateTemp && !backgroundWorkerExcelCreate.CancellationPending)
                    {
                        if (i2 % 4 == 0)
                        {
                            string filePathNoExtension = listOfExcelFiles[i2].Substring(0, listOfExcelFiles[i2].Length - 5);
                            string newFilePathCopy = filePathNoExtension + "_copy" + ".xlsx";
                            try
                            {
                                //CopyExcel(listOfExcelFiles[i2], newFilePathCopy);
                            }
                            catch (Exception ex)
                            {
                                backgroundWorkerExcelCreate.CancelAsync();
                                MessageBox.Show(ex.Message, "Error coping excels", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        i2++;
                    }*/

                    // ~~~~~~ find-replace in some excels
                    int i3 = 0;
                    while (i3 < filesToCreateTemp && !backgroundWorkerExcelCreate.CancellationPending)
                    {
                        if (i3 % 2 == 0)    // to be restored to <<i3 % 3 == 0>> (or not?)
                        {
                            try
                            {
                                workBook = excel.Workbooks.Open(listOfExcelFiles[i3], m, false, m, m, m, m, m, m, m, m, m, m, m, m);
                                workSheet = (Worksheet)workBook.ActiveSheet;

                                Range rng = (Range)workSheet.UsedRange;

                                rng.Replace(1000, textToReplace1, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, m, m, m);
                                rng.Replace(2000, textToReplace2, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, m, m, m);
                                rng.Replace(3000, textToReplace1, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, m, m, m);
                                rng.Replace(4000, textToReplace2, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, m, m, m);

                                workBook.Save();
                                workBook.Close();
                                workBook = null;
                                workSheet = null;
                                if (workSheet != null) Marshal.ReleaseComObject(workSheet);
                                if (workBook != null) Marshal.ReleaseComObject(workBook);
                            }
                            catch (Exception ex)
                            {
                                backgroundWorkerExcelCreate.CancelAsync();
                                MessageBox.Show(ex.Message, "Error in find-replace", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        i3++;
                    }


                    // ~~~~~~~~~~~~~~~~~~~ Terminating Excel instance ~~~~~~~~~~~~~~~~~~~
                    excel.Quit();
                    excel = null;
                    if (excel != null) Marshal.ReleaseComObject(excel);
                    m = null;
                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                }
                catch (Exception ex)
                {
                    backgroundWorkerExcelCreate.CancelAsync();
                    MessageBox.Show(ex.Message, "backGroundWorker ExcelCreate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                backgroundWorkerExcelCreate.Dispose();
                backgroundWorkerExcelCreate.RunWorkerCompleted -= new RunWorkerCompletedEventHandler(backgroundWorkerExcelCreate_RunWorkerCompleted);
                backgroundWorkerExcelCreate.DoWork -= new DoWorkEventHandler(backgroundWorkerExcelCreate_DoWork);
            }
        }
        // bgw EXCELCreate Completed
        private void backgroundWorkerExcelCreate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (programRunning)
            {
                timmerExcelFiles.Stop();
                TimeSpan runTimeExcelFiless = timmerExcelFiles.Elapsed;
                timmerExcelFiles.Reset();
                MessageBox.Show("Excel files created!\r\n\r\nRun time: " + runTimeExcelFiless.TotalSeconds.ToString().Substring(0, 6) + " seconds.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (!backgroundWorkerWordCreate.IsBusy)
                restoreAfterRun();
        }

        // bgw EmptyFiles creation doWork
        private void backgroundWorkerEmptyFiles_DoWork(object sender, DoWorkEventArgs e)
        {
            if (!backgroundWorkerEmptyFiles.CancellationPending)
            {
                int filesToCreateTemp = filesToCreate;             // avoid change in #files to be created during runtime
                bool usingWordFilesTemp = checkBoxUsingWord.Checked ? true : false;
                bool usingExcelFilesTemp = checkBoxUsingExcel.Checked ? true : false;

                try
                {
                    string emptyWorkingDir = workingDirectory + "officeSimulation_empty\\";
                    System.IO.Directory.CreateDirectory(emptyWorkingDir);
                    int documentCount = 0;
                    string emptyFileNameDate = emptyFileName + DateTime.Now.ToString("dd-M-yyyy_HH-mm");
                    string emptyFileNameFinal;

                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~ Word instance ~~~~~~~~~~~~~~~~~~~~~~~~~
                    var winwordEmpty = new Microsoft.Office.Interop.Word.Application();
                    winwordEmpty.ShowAnimation = false;
                    winwordEmpty.Visible = false;
                    object winWordMissing = System.Reflection.Missing.Value;
                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                    // ~~~~~~~~~~~~~~~~~~~~~~~~~ Excel instance ~~~~~~~~~~~~~~~~~~~~~~~~~
                    Microsoft.Office.Interop.Excel.Application excelEmpty;
                    Microsoft.Office.Interop.Excel.Workbook workBookEmpty;
                    Microsoft.Office.Interop.Excel.Worksheet workSheetEmpty;
                    excelEmpty = new Microsoft.Office.Interop.Excel.Application();
                    excelEmpty.Visible = false;
                    excelEmpty.DisplayAlerts = false;
                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                    while ((documentCount < filesToCreateTemp) && (!backgroundWorkerEmptyFiles.CancellationPending))
                    {
                        emptyFileNameFinal = emptyFileNameDate + "_" + documentCount;    // index of file (at filename)

                        // ~~~~ create empty word files
                        if (usingWordFilesTemp)
                        {
                            try
                            {
                                //Create a new document
                                Microsoft.Office.Interop.Word.Document emptyDocument = winwordEmpty.Documents.Add(ref winWordMissing, ref winWordMissing, ref winWordMissing, ref winWordMissing);
                                emptyDocument.Content.SetRange(0, 0);
                                //Save the document
                                object filename = emptyWorkingDir + emptyFileNameFinal + ".docx";
                                emptyDocument.SaveAs2(ref filename);
                                // Keeping a list of files created
                                listOfEmptyWordFiles.Add(filename.ToString());
                                emptyDocument.Close(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                                emptyDocument = null;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Error creating empty word file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }

                        // ~~~~ create empty excel files
                        if (usingExcelFilesTemp)
                        {
                            try
                            {
                                //Create a new WorkBook
                                workBookEmpty = excelEmpty.Workbooks.Add(Type.Missing);
                                workSheetEmpty = (Microsoft.Office.Interop.Excel.Worksheet)workBookEmpty.ActiveSheet;

                                //Save the document
                                object filename = emptyWorkingDir + emptyFileNameFinal + ".xlsx";
                                workBookEmpty.SaveAs(filename);
                                // Keeping a list of files created
                                listOfEmptyExcelFiles.Add(filename.ToString());
                                workBookEmpty.Close();
                                workBookEmpty = null;
                                workSheetEmpty = null;
                                if (workSheetEmpty != null) Marshal.ReleaseComObject(workSheetEmpty);
                                if (workBookEmpty != null) Marshal.ReleaseComObject(workBookEmpty);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Error creating empty excel file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        documentCount++;
                    }
                    // ~~~~~~~~~~~~~~~~~~~~ Terminating Word instance ~~~~~~~~~~~~~~~~~~~
                    winwordEmpty.Quit(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                    winwordEmpty.Quit();
                    if (winwordEmpty != null) Marshal.ReleaseComObject(winwordEmpty);
                    winwordEmpty = null;
                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                    // ~~~~~~~~~~~~~~~~~~~ Terminating Excel instance ~~~~~~~~~~~~~~~~~~~
                    excelEmpty.Quit();
                    excelEmpty = null;
                    if (excelEmpty != null) Marshal.ReleaseComObject(excelEmpty);
                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error creating empty files", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    backgroundWorkerEmptyFiles.Dispose();
                    backgroundWorkerEmptyFiles.RunWorkerCompleted -= new RunWorkerCompletedEventHandler(backgroundWorkerEmptyFiles_RunWorkerCompleted);
                    backgroundWorkerEmptyFiles.DoWork -= new DoWorkEventHandler(backgroundWorkerEmptyFiles_DoWork);
                }
            }
            else
            {
                backgroundWorkerEmptyFiles.Dispose();
                backgroundWorkerEmptyFiles.RunWorkerCompleted -= new RunWorkerCompletedEventHandler(backgroundWorkerEmptyFiles_RunWorkerCompleted);
                backgroundWorkerEmptyFiles.DoWork -= new DoWorkEventHandler(backgroundWorkerEmptyFiles_DoWork);
            }
        }
        // bgw EmptyFiles creation Completed
        private void backgroundWorkerEmptyFiles_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            createEmpyFilesToolStripMenuItem.BackColor = Color.FromKnownColor(KnownColor.Control);
            timmerEmptyFiles.Stop();
            TimeSpan runTimeEmptyFiles = timmerEmptyFiles.Elapsed;
            timmerEmptyFiles.Reset();
            DialogResult dialogResult = MessageBox.Show("Empty files created!\n\r\n\rPath: " + workingDirectory + "officeSimulation_empty\\" + "\r\n\r\nRun time: " + runTimeEmptyFiles.TotalSeconds.ToString().Substring(0, 6) + " seconds" + "\r\n\r\nDo you want to open containing folder?", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    System.Diagnostics.Process.Start(workingDirectory + "officeSimulation_empty\\");       // Open folder containing empty files
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error opening EmptyFiles directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                if (autoDelete)
                {
                    try
                    {
                        string deletingDir = workingDirectory + "officeSimulation_empty\\";
                        System.IO.Directory.Delete(deletingDir, true);
                        listOfEmptyWordFiles.Clear();
                        listOfEmptyExcelFiles.Clear();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error deleting \\officeSimulation_empty\\", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            if (!backgroundWorkerWordCreate.IsBusy && !backgroundWorkerExcelCreate.IsBusy)
            {
                BeginInvoke((MethodInvoker)delegate
                {
                    progressBar1.Style = ProgressBarStyle.Continuous;
                });
            }
            emptyFilesRunning = false;
        }
        
        // --- Canceling all background workers
        private void cancelBackgroundWorkers()
        {
            backgroundWorkerWordCreate.CancelAsync();
            backgroundWorkerExcelCreate.CancelAsync();
            backgroundWorkerEmptyFiles.CancelAsync();
            //MessageBox.Show("Background workers cancelled!", "Cancel bgw", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        // ---------------------------------------------------------------------------------------

        // ---------------------------------- Menu Strip options---------------------------------- 
        // About menu
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "About";
            String msgBoxAboutText = "Version 1.0  -  March 2017\r\nDeveloped by Apostolos Smyrnakis - IT/CDA/AD\r\n\r\nFor support contact: apostolos.smyrnakis@cern.ch";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Information);
        }

        // Help menu: files To Create
        private void filesToCreateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Files to create";
            String msgBoxAboutText = "Use the slider or type the number of files that the application will create.\r\nWhen using both Word and Excel files, that number will be used for .docx AND .xlsx files.";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Question);
        }

        // Help menu: working Directory
        private void workingDirectoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Working directory";
            String msgBoxAboutText = "Two subdirectories will be created inside the directory you have selected.\r\nSubdirectories names: 'officeSimulation_word' and 'officeSimulation_excel'";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Question);
        }

        // Help menu: word / excell files
        private void wordExcelFilesToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Word / Excel files";
            String msgBoxAboutText = "Select whether to perform simulations on Word and/or Excel files.\r\nAt least one option should always be selected!";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Question);
        }

        // Help menu: delay In Editing
        private void delayInEditingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Delay in editing";
            String msgBoxAboutText = "Insert a delay during documents creation. Delay is used only once per document.\r\nValue is in seconds: 0 - 3600!\r\n\r\nNote: no delay is added when creating empty word and/or excel files!";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Question);
        }

        // Help menu: auto Delete Created Files
        private void autoDeleteCreatedFilesToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Auto delete created files";
            String msgBoxAboutText = "If selected, all files created by the application will be deleted immediately after finishing operations on them.";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Question);
        }

        // Help menu: vary Formatting
        private void varyFormattingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Vary formatting";
            String msgBoxAboutText = "Turn ON/OFF formatting variations when creating files.\r\n";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Question);
        }

        // Help menu: create empty files
        private void createEmptyFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Create empty files";
            String msgBoxAboutText = "Creates empty files (word and/or excel, depending on your selection). \r\nThe number of files to be created is specified by the 'Files to create' option.\r\nFiles are saved in 'officeSimulation_empty' directory, under 'Working directory'.";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Question);
        }
        // ---------------------------------------------------------------------------------------
    }
}
