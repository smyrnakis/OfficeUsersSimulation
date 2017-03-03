/* ************ To fix: ************
 *  1) clicking STOP button -> kill winword (office instances)
 *  2) remove lorem text from code -> have it in different .h file? 
 *  3) keep all settings in memory (eg "vary formatting" etc)
 */


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
// Microsoft World Object XX library was needed and I included it! 

namespace OfficeUsersSimulation_C
{
    public partial class Form1 : Form
    {
        //public const string wordFilePath = @"\\\\cern.ch\\dfs\\Users\\a\\asmyrnak\\Documents\\Visual Studio 2015\\Projects\\EXTRA\\testWords\\";
        //public const string excelFilePath = @"\\\\cern.ch\\dfs\\Users\\a\\asmyrnak\\Documents\\Visual Studio 2015\\Projects\\EXTRA\\testExcels\\";

        // ------------------------------------ Declarations ------------------------------------
        Stopwatch runStopWatch = new Stopwatch();
        
        string workingDirectory = "";
        string sampleWordPath = "";
        string sampleExcelPath = "";
        int filesToCreate = 0;
        bool randomOrder = false;
        bool usingWordFiles = true;
        bool usingExcelFiles = false;
        bool addDelay = false;
        int addDelaySec = 0;
        bool autoDelete = true;
        bool varyFormatting = false;
        bool programRunning = false;
        string wordFileName = "testWord_";
        string excelFileName = "testExcel_";

        List<string> listOfWordFiles = new List<string>();
        List<string> listOfExcelFiles = new List<string>();

        const string lorem1 = @"Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.";
        const string lorem5 = @"Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.
Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.
Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.
Aenean nec lorem. In porttitor. Donec laoreet nonummy augue.
Suspendisse dui purus, scelerisque at, vulputate vitae, pretium mattis, nunc. Mauris eget neque at sem venenatis eleifend. Ut nonummy.
";
        const string lorem10 = @"Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.
Nunc viverra imperdiet enim.Fusce est.Vivamus a tellus.
Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas.Proin pharetra nonummy pede. Mauris et orci.
Aenean nec lorem.In porttitor. Donec laoreet nonummy augue.
Suspendisse dui purus, scelerisque at, vulputate vitae, pretium mattis, nunc.Mauris eget neque at sem venenatis eleifend.Ut nonummy.
Fusce aliquet pede non pede.Suspendisse dapibus lorem pellentesque magna.Integer nulla.
Donec blandit feugiat ligula. Donec hendrerit, felis et imperdiet euismod, purus ipsum pretium metus, in lacinia nulla nisl eget sapien.Donec ut est in lectus consequat consequat.
Etiam eget dui.Aliquam erat volutpat.Sed at lorem in nunc porta tristique.
Proin nec augue.Quisque aliquam tempor magna. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas.
Nunc ac magna.Maecenas odio dolor, vulputate vel, auctor ac, accumsan id, felis.Pellentesque cursus sagittis felis.
";
        const string lorem15 = @"Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.
Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.
Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.
Aenean nec lorem. In porttitor. Donec laoreet nonummy augue.
Suspendisse dui purus, scelerisque at, vulputate vitae, pretium mattis, nunc. Mauris eget neque at sem venenatis eleifend. Ut nonummy.
Fusce aliquet pede non pede. Suspendisse dapibus lorem pellentesque magna. Integer nulla.
Donec blandit feugiat ligula. Donec hendrerit, felis et imperdiet euismod, purus ipsum pretium metus, in lacinia nulla nisl eget sapien. Donec ut est in lectus consequat consequat.
Etiam eget dui. Aliquam erat volutpat. Sed at lorem in nunc porta tristique.
Proin nec augue. Quisque aliquam tempor magna. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas.
Nunc ac magna. Maecenas odio dolor, vulputate vel, auctor ac, accumsan id, felis. Pellentesque cursus sagittis felis.
Pellentesque porttitor, velit lacinia egestas auctor, diam eros tempus arcu, nec vulputate augue magna vel risus. Cras non magna vel ante adipiscing rhoncus. Vivamus a mi.
Morbi neque. Aliquam erat volutpat. Integer ultrices lobortis eros.
Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin semper, ante vitae sollicitudin posuere, metus quam iaculis nibh, vitae scelerisque nunc massa eget pede. Sed velit urna, interdum vel, ultricies vel, faucibus at, quam.
Donec elit est, consectetuer eget, consequat quis, tempus quis, wisi. In in nunc. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos hymenaeos.
Donec ullamcorper fringilla eros. Fusce in sapien eu purus dapibus commodo. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus.";
        const string lorem20 = @"Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.
Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.
Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.
Aenean nec lorem. In porttitor. Donec laoreet nonummy augue.
Suspendisse dui purus, scelerisque at, vulputate vitae, pretium mattis, nunc. Mauris eget neque at sem venenatis eleifend. Ut nonummy.
Fusce aliquet pede non pede. Suspendisse dapibus lorem pellentesque magna. Integer nulla.
Donec blandit feugiat ligula. Donec hendrerit, felis et imperdiet euismod, purus ipsum pretium metus, in lacinia nulla nisl eget sapien. Donec ut est in lectus consequat consequat.
Etiam eget dui. Aliquam erat volutpat. Sed at lorem in nunc porta tristique.
Proin nec augue. Quisque aliquam tempor magna. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas.
Nunc ac magna. Maecenas odio dolor, vulputate vel, auctor ac, accumsan id, felis. Pellentesque cursus sagittis felis.
Pellentesque porttitor, velit lacinia egestas auctor, diam eros tempus arcu, nec vulputate augue magna vel risus. Cras non magna vel ante adipiscing rhoncus. Vivamus a mi.
Morbi neque. Aliquam erat volutpat. Integer ultrices lobortis eros.
Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin semper, ante vitae sollicitudin posuere, metus quam iaculis nibh, vitae scelerisque nunc massa eget pede. Sed velit urna, interdum vel, ultricies vel, faucibus at, quam.
Donec elit est, consectetuer eget, consequat quis, tempus quis, wisi. In in nunc. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos hymenaeos.
Donec ullamcorper fringilla eros. Fusce in sapien eu purus dapibus commodo. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus.
Cras faucibus condimentum odio. Sed ac ligula. Aliquam at eros.
Etiam at ligula et tellus ullamcorper ultrices. In fermentum, lorem non cursus porttitor, diam urna accumsan lacus, sed interdum wisi nibh nec nisl. Ut tincidunt volutpat urna.
Mauris eleifend nulla eget mauris. Sed cursus quam id felis. Curabitur posuere quam vel nibh.
Cras dapibus dapibus nisl. Vestibulum quis dolor a felis congue vehicula. Maecenas pede purus, tristique ac, tempus eget, egestas quis, mauris.
Curabitur non eros. Nullam hendrerit bibendum justo. Fusce iaculis, est quis lacinia pretium, pede metus molestie lacus, at gravida wisi ante at libero.";

        string[] lorems = {lorem1, lorem5, lorem10, lorem15, lorem20};
        // ---------------------------------------------------------------------------------------

        // ------------ Initialize values & items - Restore last settings from memory ------------
        public Form1()
        {
            //Create an instance for word app
            var winword = new Word.Application();
            InitializeComponent();
            textBox2.Text = Properties.Settings.Default["lastWorkingDir"].ToString();   // Load last working directory
            workingDirectory = @textBox2.Text;
            textBox3.Text = Properties.Settings.Default["lastSampleWord"].ToString();   // Load last sample word path
            sampleWordPath = textBox3.Text;
            textBox4.Text = Properties.Settings.Default["lastSampleExcel"].ToString();  // load last sample excel path
            sampleExcelPath = textBox4.Text;
            trackBar1.Value = Properties.Settings.Default.lastSliderValue;              // load last number of created files
            filesToCreate = Properties.Settings.Default.lastSliderValue;
            numericUpDown2.Value = trackBar1.Value;
            checkBox6.Checked = Properties.Settings.Default.lastVaryFormatting;         // load last vary-formating option
            checkBox5.Checked = Properties.Settings.Default.lastAutoDelete;             // load last auto-delete option
            radioButton1.Select();              // Default: sequentially
            checkBox1.Checked = true;           // Default: Word files
            textBox1.Visible = false;           // Hide textBox1 (used for logging durring program running)
            this.FormBorderStyle = FormBorderStyle.Fixed3D;
            ActiveControl = trackBar1;          // This should be always the last one, after setting all other properties!
        }
        // ---------------------------------------------------------------------------------------

        // -------------------- Track changes in "number of files" to create ---------------------
        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            trackBar1.Value = Convert.ToInt32(numericUpDown2.Value);
            filesToCreate = Convert.ToInt32(numericUpDown2.Value);
            Properties.Settings.Default["lastSliderValue"] = trackBar1.Value;
            Properties.Settings.Default.Save();
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            numericUpDown2.Value = trackBar1.Value;
            filesToCreate = trackBar1.Value;
            Properties.Settings.Default["lastSliderValue"] = trackBar1.Value;
            Properties.Settings.Default.Save();
        }
        // ---------------------------------------------------------------------------------------

        // -------------------------- Restore default button appearance -------------------------- 
        public void restoreAfterRun()
        {
            programRunning = false;
            textBox1.Clear();
            textBox1.Visible = false;
            progressBar1.Value = 0;
            progressBar1.Update();
            button1.Text = "R U N";
            button1.ForeColor = System.Drawing.Color.ForestGreen;
            if (autoDelete)
            {
                if (usingWordFiles)
                {
                    try
                    {
                        string deletingDir = workingDirectory + "\\Word Files\\";
                        System.IO.Directory.Delete(deletingDir, true);
                        listOfWordFiles.Clear();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                if (usingExcelFiles)
                {
                    try
                    {
                        string deletingDir = workingDirectory + "\\Excel Files\\";
                        System.IO.Directory.Delete(deletingDir, true);
                        listOfExcelFiles.Clear();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        // ---------------------------------------------------------------------------------------
        
        // ------------------------ Message Box with all variable values ------------------------- 
        private void debuggingMsgBox()
        {
            String msgBoxAboutCaption = "Debugging messageBox";
            String msgBoxAboutText = "filesToCreate: " + filesToCreate + "\r\n\r\nworkingDirectory: " + workingDirectory + "\r\n\r\nsampleWordPath: " + sampleWordPath + "\r\n\r\nsampleExcelPath: " + sampleExcelPath + "\r\n\r\nrandomOrder: " + randomOrder + "\r\n\r\nusingWordFiles: " + usingWordFiles + "\r\nusingExcelFiles: " + usingExcelFiles + "\r\n\r\naddDelay: " + addDelay + "\r\naddDelaySec: " + addDelaySec + "\r\n\r\nautoStop: " + "\r\n\r\nautoDelete: " + autoDelete + "\r\nvaryFormatting: " + varyFormatting;
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons);
        }
        // ---------------------------------------------------------------------------------------

        // Button RUN click  
        private void button1_Click(object sender, EventArgs e)
        {
            if (!checkBox1.Checked && !checkBox2.Checked)
            {
                MessageBox.Show("Please select at least one between 'Word' or 'Excel' files!");
            }
            else if (textBox2.Text.Length < 3)
            {
                MessageBox.Show("Please select a working directory!");
            }
            else if (checkBox1.Checked && textBox3.Text.Length < 3)
            {
                MessageBox.Show("Please select sample word file!");
            }
            else if (checkBox2.Checked && textBox4.Text.Length < 3)
            {
                MessageBox.Show("Please select sample excel file!");
            }
            else if (programRunning)        // canceling all background workers & restoring defaults
            {
                backgroundWorkerWordCreate.CancelAsync();
                backgroundWorkerExcelCreate.CancelAsync();
                //backgroundWorker~~~~~~~~.CancelAsync();
                //backgroundWorker~~~~~~~~.CancelAsync();
                if (checkBox5.Checked)
                {
                    //Delete created files!
                }
                Thread.Sleep(100);
                restoreAfterRun();
                Thread.Sleep(100);
                MessageBox.Show("Program terminated by the user!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                restoreAfterRun();
            }
            else if (!programRunning)
            {
                try
                {
                    programRunning = true;
                    //runStopWatch.Start();
                    button1.ForeColor = System.Drawing.Color.Red;
                    button1.Text = "S T O P";
                    textBox1.Visible = true;
                    textBox1.BringToFront();

                    //debuggingMsgBox();

                    if (checkBox1.Checked)
                        backgroundWorkerWordCreate.RunWorkerAsync();

                    if (checkBox2.Checked)
                        backgroundWorkerExcelCreate.RunWorkerAsync();
                    
                    /*
                    string fromWordPath = workingDirectory + "\\" + wordFileName + "_from.docx";
                    string toWordPath = workingDirectory + "\\" + wordFileName + "_to.docx";
                    
                    CopyDocument(fromWordPath, toWordPath);
                    runStopWatch.Stop();
                    TimeSpan runSWtimeSpan = runStopWatch.Elapsed;
                    MessageBox.Show(runSWtimeSpan.ToString(),"Run time");
                    */

                }
                catch (Exception ex)
                {
                    restoreAfterRun();
                    // to do: .CancelAsync(); all background workers
                    backgroundWorkerWordCreate.CancelAsync();
                    backgroundWorkerExcelCreate.CancelAsync();
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                restoreAfterRun();
                // to do: .CancelAsync(); all background workers
                backgroundWorkerWordCreate.CancelAsync();
                backgroundWorkerExcelCreate.CancelAsync();
                MessageBox.Show("Error! 'else' case executed in button1_Click!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ------------------- Create document method ------------------- 
        private void CreateDocument(string filePath, string fileName)
        {
            //Create an instance for word app
            var winword = new Word.Application();
            //Set animation status for word application
            winword.ShowAnimation = false;
            //Set status for word application is to be visible or not.
            winword.Visible = false;

            //Create a missing variable for missing value
            object winWordMissing = System.Reflection.Missing.Value;

            if (varyFormatting)
            {
                try
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
                    // adding random delay
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
                    object filename = filePath + "\\" + fileName + ".docx";
                    document.SaveAs2(ref filename);
                    // Keeping a list of files created
                    listOfWordFiles.Add(filename.ToString());
                    document.Close(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                    document = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                try
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
                    // adding random delay
                    if (addDelay)
                        Thread.Sleep(addDelaySec * 1000);
                    document.Content.Text = lorems[rLorem2] + "\r\n";
                    document.Content.Text = lorems[rLorem3];
                    
                    //Save the document
                    object filename = filePath + "\\" + fileName + ".docx";
                    document.SaveAs2(ref filename);
                    // Keeping a list of files created
                    listOfWordFiles.Add(filename.ToString());
                    document.Close(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                    document = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                winword.Quit(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                winword = null;
            }
        }

        // ------------------- Copy document method ------------------- 
        private void CopyDocument(string fromWordPath, string toWordPath)
        {
            var application = new Word.Application();                           // Start MS Word application
            var originalDocument = application.Documents.Open(fromWordPath);    // Open original document

            originalDocument.ActiveWindow.Selection.WholeStory();               // Select all in original document
            var originalText = originalDocument.ActiveWindow.Selection;         // Copy everything to the variable

            var newDocument = new Word.Document();                              // Create new Word document
            newDocument.Range().Text = originalText.Text;                       // Pasete everything from the variable
            newDocument.SaveAs(toWordPath); // maybe SaveAs2??                  // Save the new document

            originalDocument.Close(false);
            newDocument.Close();

            application.Quit();
            application = null;

            restoreAfterRun();              // Restoring default appearance
        }

       
        // -------------------------- Working Directory - Sample Files -------------------------- 
        // workingDirectory
        private void button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox2.Text = folderBrowserDialog1.SelectedPath;
                    workingDirectory = @textBox2.Text;
                    Properties.Settings.Default["lastWorkingDir"] = workingDirectory;
                    Properties.Settings.Default.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            workingDirectory = @textBox2.Text;
            Properties.Settings.Default["lastWorkingDir"] = workingDirectory;
            Properties.Settings.Default.Save();
        }

        // Sample Word file directory
        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox3.Text = openFileDialog1.FileName;
                    sampleWordPath = textBox3.Text;                                     // sampleWordPath = @textBox3.Text; 
                    Properties.Settings.Default["lastSampleWord"] = textBox3.Text;
                    Properties.Settings.Default.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            sampleWordPath = textBox3.Text;                                     // sampleWordPath = @textBox3.Text; 
            Properties.Settings.Default["lastSampleWord"] = textBox3.Text;
            Properties.Settings.Default.Save();
        }

        // Sample Excel file directory
        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox4.Text = openFileDialog2.FileName;
                    sampleExcelPath = textBox4.Text;                                    // sampleExcelPath = @textBox4.Text;
                    Properties.Settings.Default["lastSampleExcel"] = textBox4.Text;
                    Properties.Settings.Default.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            sampleExcelPath = textBox4.Text;                                    // sampleExcelPath = @textBox4.Text;
            Properties.Settings.Default["lastSampleExcel"] = textBox4.Text;
            Properties.Settings.Default.Save();
        }
        // ---------------------------------------------------------------------------------------

        // ---------------------------------- Menu Strip options---------------------------------- 
        // About menu
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "About";
            String msgBoxAboutText = "Version 1.0\r\nDeveloped by Apostolos Smyrnakis - IT/CDA/AD\r\n\r\nFor support contact: apostolos.smyrnakis@cern.ch";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Information);
        }

        // Help menu: files To Create
        private void filesToCreateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Files to create";
            String msgBoxAboutText = "Use the slider or type the number of files that the application will create.\r\nWhen using both Word and Excel files, that number will be used for .docx AND .xlsx files.\r\nTotal files to be created: (Word * selection) + (Excel * selection)";
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

        // Help menu: load Samle Files - directories
        private void loadSampleWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Loading sample files";
            String msgBoxAboutText = "Select the default .docx & .xlsx files provided with the application\r\nor use your own files.\r\nSample files must contain Lorem Ipsum words!";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Question);
        }

        // Help menu: sequentially
        private void sequentiallyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Sequentially";
            String msgBoxAboutText = "Operations order: 'Sequentially'\r\n\r\n1) Create files.\r\n2) Open one by one and copy sample data inside.\r\n3) Open one by one and replace some data.\r\n4) Open one by one and append some text\r\n    (using =rand() & =lorem() MS functions).";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Question);
        }

        // Help menu: randomly
        private void randomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Randomly";
            String msgBoxAboutText = "Operations order: 'Randomly'\r\n\r\n1) Create files.\r\n2) Open one by one and copy sample data inside.\r\n3) Open randomly 1/2 of files and replace some data.\r\n4) Delete 1/4 of files.\r\n5) Create new files (same quantity with the ones deleted).\r\n6) Open one by one and write some data\r\n    (using =rand() & =lorem() MS functions).\r\n7) Open randomly 1/2 of files and find-replace some words.";
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
            String msgBoxAboutText = "Insert a delay during word document creation. Delay is used only once per document.\r\nValue is in seconds: 0 - 3600!";
            MessageBoxButtons msgAboutButtons = MessageBoxButtons.OK;
            DialogResult result;
            result = MessageBox.Show(msgBoxAboutText, msgBoxAboutCaption, msgAboutButtons, MessageBoxIcon.Question);
        }
        
        // Help menu: auto Delete Created Files
        private void autoDeleteCreatedFilesToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            String msgBoxAboutCaption = "Auto delete created files";
            String msgBoxAboutText = "Select whether to delete all files created by the application upon completion or keep them for debugging purposes.";
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
        // ---------------------------------------------------------------------------------------

        // --------------------------------- Check Boxes settings -------------------------------- 
        // checkBox: use Word files 
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            usingWordFiles = checkBox1.Checked ? true : false;
        }

        // checkBox: use Excel files
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            usingExcelFiles = checkBox2.Checked ? true : false;
        }

        // checkBox: add some seconds delay during document edit
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            addDelay = checkBox3.Checked ? true : false;
            addDelaySec = Convert.ToInt32(numericUpDown1.Value);
            if (!addDelay)
                numericUpDown1.Value = 0;
        }

        // checkBox: auto delete created files
        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            autoDelete = checkBox5.Checked ? true : false;
            Properties.Settings.Default["lastAutoDelete"] = checkBox5.Checked;
            Properties.Settings.Default.Save();
        }

        // checkBox: vary documents formatting
        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            varyFormatting = checkBox6.Checked ? true : false;
            Properties.Settings.Default["lastVaryFormatting"] = checkBox6.Checked;
            Properties.Settings.Default.Save();
        }
        // ---------------------------------------------------------------------------------------

        // ---------------------------------- Other items check ---------------------------------- 
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            checkBox3.Checked = true;
            addDelaySec = Convert.ToInt32(numericUpDown1.Value);
            if (numericUpDown1.Value == 0)
                checkBox3.Checked = addDelay = false;
        }
        
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            randomOrder = false;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            randomOrder = true;
        }
        // ---------------------------------------------------------------------------------------

      
        // -------------------------------- Program exit methods --------------------------------- 
/*
        private void OnApplicationExit(object sender, EventArgs e)
        {
            Console.WriteLine("Exiting...");
            MessageBox.Show("Exiting application...");
            if (checkBox5.Checked)
            {
                // Delete created files
            }
        }

        private void HMI_FormClosing(object sender, FormClosingEventArgs e)
        {
            Console.WriteLine("Exiting...");
            MessageBox.Show("Exiting application...");
            if (checkBox5.Checked)
            {
                // Delete created files
            }
        }
        // ---------------------------------------------------------------------------------------
*/

        // ----------------------------- Background workers methods ------------------------------
        // bgw WordCreate Completed
        private void backgroundWorkerWordCreate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (programRunning)
                MessageBox.Show("Word files created!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (!backgroundWorkerExcelCreate.IsBusy)
                restoreAfterRun();
            MessageBox.Show(listOfWordFiles[0] + "\r\n\r\n" + listOfWordFiles[1] + "\r\n\r\n" + listOfWordFiles[2] + "\r\n\r\n" + listOfWordFiles[3], "Files created", MessageBoxButtons.OK, MessageBoxIcon.None);
        /*
            for (int i = 0; i < listOfWordFiles.Count; i++)
            {
                MessageBox.Show(listOfWordFiles[0] + listOfWordFiles[1] + listOfWordFiles[2] + listOfWordFiles[3], "Files created", MessageBoxButtons.OK, MessageBoxIcon.None);
            } */
        }
        // bgw ExcelCreate Completed
        private void backgroundWorkerExcelCreate_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (programRunning)
                MessageBox.Show("Excel files created!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (!backgroundWorkerWordCreate.IsBusy)
                restoreAfterRun();
        }
        // bgw WordCreate Progress
        private void backgroundWorkerWordCreate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            progressBar1.Update();
            textBox1.Text = String.Format("Creating word files... {0}%", e.ProgressPercentage);
        }
        // bgw ExcelCreate Progress
        private void backgroundWorkerExcelCreate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //progressBar1.Value = e.ProgressPercentage;
            //progressBar1.Update();
            //textBox1.Text = String.Format("Creating excel files... {0}%", e.ProgressPercentage);
        }
        // bgw WordCreate doWork
        private void backgroundWorkerWordCreate_DoWork(object sender, DoWorkEventArgs e)
        {
            int documentCountIndex = 1;
            try
            {
                string wordFileNameDate = wordFileName + DateTime.Now.ToString("dd-MM-yyyy_HH-mm");
                string workingDir = workingDirectory + "\\Word Files\\";
                System.IO.Directory.CreateDirectory(workingDir);
                int documentCount = 0;
                while (documentCount < filesToCreate)
                {
                    if (!backgroundWorkerWordCreate.CancellationPending)
                    {
                        string wordFileNameDateIndex = wordFileNameDate + "_" + documentCount;
                        
                        //var winword = new Word.Application();
                        CreateDocument(workingDir, wordFileNameDateIndex);
                        //winword.Quit(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                        //winword.Quit();
                        //winword = null;

                        backgroundWorkerWordCreate.ReportProgress(documentCountIndex++ * 100 / filesToCreate, string.Format("Process data {0}", documentCount));
                        documentCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                backgroundWorkerWordCreate.CancelAsync();
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        
        // ---------------------------------------------------------------------------------------
    }
}
