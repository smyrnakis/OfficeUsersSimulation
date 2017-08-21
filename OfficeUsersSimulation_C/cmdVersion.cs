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
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsApplication1;
using OfficeUsersSimulation_C;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Drawing;

namespace OfficeUsersSimulation_C
{
    public class cmdVersion
    {
        // ------------------------------------ Declarations ------------------------------------
        private Form1 frm1 = new Form1();
        public Form1 getForm1()
        {
            return frm1;
        }

        public bool argumentsOk = true;
        public bool createEmptyFilesCmd = false;
        public bool createFilesCmd = false;
        public int pathsResultCmd = -1;
        public bool tskCreateEmptyRunning = false;
        public bool tskCreateWordsRunning = false;
        public bool tskCreateExcelsRunning = false;
        //public bool tskCopyWordsRunning = false;

        // ---------------------------------------------------------------------------------------

        // -------------------------- Check validity of input arguments --------------------------
        public void checkDataValidity()
        {
            if (!createFilesCmd && !createEmptyFilesCmd)
            {
                argumentsOk = false;
                Console.WriteLine();
                Console.WriteLine("Error!");
                Console.WriteLine("Please select at least one of either 'crtfls' or 'emtfls' options!");
                Console.WriteLine("crtfls : create word and/or excel files with text inside");
                Console.WriteLine("emtfls : create empty word and/or excel files");
                Console.WriteLine();
            }

            if (frm1.filesToCreate < 4 || frm1.filesToCreate > 1000)
            {
                argumentsOk = false;
                Console.WriteLine();
                Console.WriteLine("Error!");
                Console.WriteLine("'Number of files' to be created should be [4 to 1000]!");
                Console.WriteLine("nof + <integer[4-1000]> : number of files to create");
                Console.WriteLine();
            }
            pathsResultCmd = frm1.checkLoadedPaths();
            switch (pathsResultCmd)
            {
                // 0: workingDirectory EXISTS , sampleFiles EXIST
                case 0:
                    break;
                // -1: workingDirectory NOT EXISTS , sampleFiles NOT EXIST
                case -1:
                    argumentsOk = false;
                    Console.WriteLine();
                    Console.WriteLine("Error!");
                    Console.WriteLine("Please check the working path and the sample files!");
                    Console.WriteLine("wrkdir + <path> : working directory");
                    Console.WriteLine("smwd + <path.docx> : sample word file path");
                    Console.WriteLine("smxl + <path.xlsx> : sample excel file path");
                    Console.WriteLine();
                    Console.WriteLine("Selected working directory: " + frm1.workingDirectory);
                    //Console.WriteLine("Selected sample word directory: " + frm1.sampleWordPath);
                    //Console.WriteLine("Selected sample excel directory: " + frm1.sampleExcelPath);
                    Console.WriteLine();
                    break;
                // 1: workingDirectory EXISTS , sampleFiles NOT EXIST
                case 1:
                    if (createFilesCmd)
                    {
                        argumentsOk = false;
                        Console.WriteLine();
                        Console.WriteLine("Warning!");
                        Console.WriteLine("Please check the sample files paths!");
                        Console.WriteLine("smwd + <path.docx> : sample word file path");
                        Console.WriteLine("smxl + <path.xlsx> : sample excel file path");
                        Console.WriteLine();
                        //Console.WriteLine("Selected sample word directory: " + frm1.sampleWordPath);
                        //Console.WriteLine("Selected sample excel directory: " + frm1.sampleExcelPath);
                        Console.WriteLine();
                    }
                    break;
                // 2: workingDirectory NOT EXISTS , sampleFiles EXIST
                case 2:
                    argumentsOk = false;
                    Console.WriteLine();
                    Console.WriteLine("Error!");
                    Console.WriteLine("Please check the working path!");
                    Console.WriteLine("wrkdir + <path> : working directory");
                    Console.WriteLine();
                    Console.WriteLine("Selected working directory: " + frm1.workingDirectory);
                    Console.WriteLine();
                    break;
                // Normaly, should never enter here! 
                default:
                    argumentsOk = false;
                    Console.WriteLine();
                    Console.WriteLine("Error!");
                    Console.WriteLine("Wrong variable 'pathsResultCmd'");
                    Console.WriteLine();
                    break;
            }
            if (!frm1.usingWordFiles && !frm1.usingExcelFiles)
            {
                argumentsOk = false;
                Console.WriteLine();
                Console.WriteLine("Error!");
                Console.WriteLine("Select at least one of the 'usewd' or 'usexl' options!");
                Console.WriteLine("usewd : using word files");
                Console.WriteLine("usexl : using excel files");
                Console.WriteLine();
            }
            if (frm1.addDelaySec < 0 || frm1.addDelaySec > 3600)
            {
                argumentsOk = false;
                Console.WriteLine();
                Console.WriteLine("Error!");
                Console.WriteLine("Delay should be 0 - 3600 seconds!");
                Console.WriteLine("delaysc + <integer[0-3600]> : delay in seconds during documents creation");
                Console.WriteLine();
            }
        }
        // ---------------------------------------------------------------------------------------

        // -------------------------------- Print variables data ---------------------------------
        public void printData()
        {
            Console.WriteLine();
            Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~ Input arguments overview ~~~~~~~~~~~~~~~~~~~~~~~~~");
            Console.WriteLine();
            Console.WriteLine("nof: " + frm1.filesToCreate);
            Console.WriteLine("emtfls: " + createEmptyFilesCmd);
            Console.WriteLine("crtfls: " + createFilesCmd);
            Console.WriteLine("usewd: " + frm1.usingWordFiles);
            Console.WriteLine("usexl: " + frm1.usingExcelFiles);
            Console.WriteLine("vrfrm: " + frm1.varyFormatting);
            Console.WriteLine("delaysc: " + frm1.addDelay + ": " + frm1.addDelaySec + " sec");
            Console.WriteLine("autodel: " + frm1.autoDelete);
            Console.WriteLine("wrkdir: " + frm1.workingDirectory);
            //Console.WriteLine("smwd: " + frm1.sampleWordPath);
            //Console.WriteLine("smxl: " + frm1.sampleExcelPath);
            Console.WriteLine();
            Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
            Console.WriteLine();
            Console.WriteLine();
        }
        // ---------------------------------------------------------------------------------------

        public void saveAllParameters()
        {
            Properties.Settings.Default.lastSliderValue = frm1.filesToCreate;
            Properties.Settings.Default["lastWorkingDir"] = frm1.workingDirectory;
            //Properties.Settings.Default["lastSampleWord"] = frm1.sampleWordPath;
            //Properties.Settings.Default["lastSampleExcel"] = frm1.sampleExcelPath;
            Properties.Settings.Default.useWords = frm1.usingWordFiles;
            Properties.Settings.Default.useExcels = frm1.usingExcelFiles;
            Properties.Settings.Default.lastVaryFormatting = frm1.varyFormatting;
            Properties.Settings.Default.lastAutoDelete = frm1.autoDelete;
            Properties.Settings.Default.lastDelayInEditing = frm1.addDelaySec;
            Properties.Settings.Default.Save();
        }

        // ---------------------------------- Run all procedures --------------------------------- 
        public void runEverything()
        {
            if (createEmptyFilesCmd)
            {
                Console.WriteLine();
                Console.WriteLine("Creating " + frm1.filesToCreate + " empty files ..........");
                Console.WriteLine();
                tskCreateEmptyRunning = true;
                Task emptyFilesCreationTSK = new Task(() => createEmptyFiles());
                emptyFilesCreationTSK.Start();
            }

            if (createFilesCmd)
            {
                if (frm1.usingWordFiles)
                {
                    Console.WriteLine();
                    Console.WriteLine("Creating " + frm1.filesToCreate + " word files ..........");
                    Console.WriteLine();
                    tskCreateWordsRunning = true;
                    Task wordsCreationTSK = new Task(() => createWords());
                    wordsCreationTSK.Start();
                }
                if (frm1.usingExcelFiles)
                {
                    Console.WriteLine();
                    Console.WriteLine("Creating " + frm1.filesToCreate + " excel files ..........");
                    Console.WriteLine();
                    tskCreateExcelsRunning = true;
                    Task excelsCreationTSK = new Task(() => createExcels());
                    excelsCreationTSK.Start();
                }
            }
            
            if (frm1.autoDelete)
            {
                Console.WriteLine();
                Console.WriteLine("Deleting \\officeSimulation_word\\  &  \\officeSimulation_excel\\ directories ..........");
                Console.WriteLine();
                frm1.autoDeleteHandler();
                Console.WriteLine("runEverything(): frm1.autoDeleteHandler() -> DONE");
            }
        }
        // ---------------------------------------------------------------------------------------
        public void waitAllToFinish()
        {
            Console.WriteLine();
            
            do
            {
                Console.Write(" . ");
                Thread.Sleep(1500);
            } while (tskCreateEmptyRunning || tskCreateWordsRunning || tskCreateExcelsRunning);
            Console.WriteLine();
            Console.WriteLine("waitAllToFinish() -> DONE");
        }

        private void createEmptyFiles()
        {
            int filesToCreateTemp = frm1.filesToCreate;
            try
            {
                string emptyWorkingDir = frm1.workingDirectory + "Empty Files\\";
                System.IO.Directory.CreateDirectory(emptyWorkingDir);
                int documentCount = 0;
                string emptyFileNameDate = frm1.emptyFileName + DateTime.Now.ToString("dd-M-yyyy_HH-mm");
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

                while (documentCount < filesToCreateTemp)
                {
                    emptyFileNameFinal = emptyFileNameDate + "_" + documentCount;    // index of file (at filename)

                    // ~~~~ create empty word files
                    if (frm1.usingWordFiles)
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
                            frm1.listOfEmptyWordFiles.Add(filename.ToString());
                            emptyDocument.Close(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                            emptyDocument = null;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error creating empty word file");
                            Console.WriteLine(ex.Message);
                        }
                    }
                    
                    // ~~~~ create empty excel files
                    if (frm1.usingExcelFiles)
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
                            frm1.listOfEmptyExcelFiles.Add(filename.ToString());
                            workBookEmpty.Close();
                            workBookEmpty = null;
                            workSheetEmpty = null;
                            if (workSheetEmpty != null) Marshal.ReleaseComObject(workSheetEmpty);
                            if (workBookEmpty != null) Marshal.ReleaseComObject(workBookEmpty);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error creating empty excel file");
                            Console.WriteLine(ex.Message);
                        }
                        //workBookEmpty = null;
                        //workSheetEmpty = null;
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
                Console.WriteLine("Error creating empty files");
                Console.WriteLine(ex.Message);
            }
            tskCreateEmptyRunning = false;
            Console.WriteLine();
            Console.WriteLine("runEverything(): createEmptyFiles() -> DONE");
        }

        private void createWords()
        {
            int filesToCreateTemp = frm1.filesToCreate;             // avoid change in '#files to be created' during runtime
            try
            {
                string workingDir = frm1.workingDirectory + "officeSimulation_word\\";
                System.IO.Directory.CreateDirectory(workingDir);
                int documentCount = 0;
                string wordFileNameDate = frm1.wordFileName + DateTime.Now.ToString("dd-MM-yyyy_HH-mm");
                string wordFileNameDateFinal;

                // ~~~~~~~~~~~~~~~~~~~~~~~~~~ Word instance ~~~~~~~~~~~~~~~~~~~~~~~~~
                var winword = new Microsoft.Office.Interop.Word.Application();
                winword.ShowAnimation = false;
                winword.Visible = false;
                object winWordMissing = System.Reflection.Missing.Value;
                // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                while (documentCount < filesToCreateTemp)
                {
                    wordFileNameDateFinal = wordFileNameDate + "_" + documentCount;
                    // ~~~~~~ create word files (vary formatting)
                    try
                    {
                        if (frm1.varyFormatting)
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
                                headerRange.Text = frm1.lorems[rLorem4];
                            }

                            //Add the footers into the document
                            foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                            {
                                //Get the footer range and add the footer details.
                                Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                                footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                                footerRange.Font.Size = 10;
                                footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                footerRange.Text = frm1.lorems[rLorem5];
                            }

                            //adding text to document
                            document.Content.SetRange(0, 0);
                            document.Content.Text = frm1.lorems[rLorem1] + "\r\n";
                            // adding some delay
                            if (frm1.addDelay)
                                Thread.Sleep(frm1.addDelaySec * 1000);
                            document.Content.Text = frm1.lorems[rLorem2] + "\r\n";
                            document.Content.Text = frm1.lorems[rLorem3];

                            //Add paragraph with Heading 1 style
                            Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref winWordMissing);
                            object styleHeading1 = "Heading 1";
                            para1.Range.set_Style(ref styleHeading1);
                            para1.Range.Text = frm1.lorems[rLorem1];
                            para1.Range.InsertParagraphAfter();

                            //Add paragraph with Heading 2 style
                            Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref winWordMissing);
                            object styleHeading2 = "Heading 2";
                            para2.Range.set_Style(ref styleHeading2);
                            para2.Range.Text = frm1.lorems[rLorem2];
                            para2.Range.InsertParagraphAfter();

                            //Save the document
                            object filename = workingDir + wordFileNameDateFinal + ".docx";
                            //object filename = filePath + "\\" + fileName + ".docx";
                            document.SaveAs2(ref filename);
                            // Keeping a list of files created
                            frm1.listOfWordFiles.Add(filename.ToString());
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
                            document.Content.Text = frm1.lorems[rLorem1] + "\r\n";
                            // adding some delay
                            if (frm1.addDelay)
                                Thread.Sleep(frm1.addDelaySec * 1000);
                            document.Content.Text = frm1.lorems[rLorem2] + "\r\n";
                            document.Content.Text = frm1.lorems[rLorem3];

                            //Save the document
                            object filename = workingDir + wordFileNameDateFinal + ".docx";
                            //object filename = filePath + "\\" + fileName + ".docx";
                            document.SaveAs2(ref filename);
                            // Keeping a list of files created
                            frm1.listOfWordFiles.Add(filename.ToString());
                            document.Close(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                            document = null;
                            if (document != null) Marshal.ReleaseComObject(document);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error creating Word file!");
                        Console.WriteLine(ex.Message);
                    }
                    documentCount++;
                }

                // ~~~~~~ copy some of the documents
                int i2 = 0;
                while (i2 < filesToCreateTemp)
                {
                    if (i2 % 4 == 0)
                    {
                        string filePathNoExtension = frm1.listOfWordFiles[i2].Substring(0, frm1.listOfWordFiles[i2].Length - 5);
                        string newFilePathCopy = filePathNoExtension + "_copy" + ".docx";
                        try
                        {
                            var originalDocument = winword.Documents.Open(frm1.listOfWordFiles[i2]);    // Open original document

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
                            Console.WriteLine("Error coping documents");
                            Console.WriteLine(ex.Message);
                        }
                    }
                    i2++;
                }

                // ~~~~~~ find-replace in some documents
                int i3 = 0;
                while (i3 < filesToCreateTemp)
                {
                    if (i3 % 3 == 0)
                    {
                        try
                        {
                            Microsoft.Office.Interop.Word.Document aDoc = winword.Documents.Open(frm1.listOfWordFiles[i3], ReadOnly: false, Visible: false);
                            aDoc.Activate();

                            winword.Selection.Find.Execute(frm1.textToSearch1, false, true, false, false, false, true, 1, false, frm1.textToReplace1, 2, false, false, false, false);
                            winword.Selection.Find.Execute(frm1.textToSearch2, false, true, false, false, false, true, 1, false, frm1.textToReplace2, 2, false, false, false, false);

                            aDoc.SaveAs2();

                            aDoc.Close(ref winWordMissing, ref winWordMissing, ref winWordMissing);
                            aDoc = null;
                            if (aDoc != null) Marshal.ReleaseComObject(aDoc);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error in find-replace");
                            Console.WriteLine(ex.Message);
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
                Console.WriteLine("backGroundWorker WordCreate Error");
                Console.WriteLine(ex.Message);
            }

            tskCreateWordsRunning = false;
            Console.WriteLine();
            Console.WriteLine("runEverything(): createWords() -> DONE");
        }

        private void createExcels()
        {
            int filesToCreateTemp = frm1.filesToCreate;             // avoid change in #files to be created during runtime
            try
            {
                string workingDir = frm1.workingDirectory + "officeSimulation_excel\\";
                System.IO.Directory.CreateDirectory(workingDir);
                int excelCount = 0;
                string excelFileNameDate = frm1.excelFileName + DateTime.Now.ToString("dd-MM-yyyy_HH-mm");
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

                while (excelCount < filesToCreateTemp)
                {
                    //Selecting random rows and columns count
                    Random r = new Random(DateTime.Now.Millisecond);
                    int rowsToWrite = r.Next(10, 70);
                    int colsToWrite = r.Next(10, 30);

                    excelFileNameDateFinal = excelFileNameDate + "_" + excelCount;
                    try
                    {
                        // ~~~~~~ create excel files (vary formatting)
                        if (frm1.varyFormatting)
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
                                    if (frm1.addDelay)
                                        Thread.Sleep(frm1.addDelaySec * 1000);   // delay in editing
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
                                    if (frm1.addDelay)
                                        Thread.Sleep(frm1.addDelaySec * 1000);
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
                        frm1.listOfExcelFiles.Add(filename.ToString());
                        workBook.Close();
                        workBook = null;
                        workSheet = null;
                        if (workSheet != null) Marshal.ReleaseComObject(workSheet);
                        if (workBook != null) Marshal.ReleaseComObject(workBook);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error creating Excel file!");
                        Console.WriteLine(ex.Message);
                    }
                    excelCount++;
                }

                // ~~~~~~ copy some of the excels
                /*int i2 = 0;
                while (i2 < filesToCreateTemp)
                {
                    if (i2 % 4 == 0)
                    {
                        string filePathNoExtension = frm1.listOfExcelFiles[i2].Substring(0, frm1.listOfExcelFiles[i2].Length - 5);
                        string newFilePathCopy = filePathNoExtension + "_copy" + ".xlsx";
                        try
                        {
                            //CopyExcel(frm1.listOfExcelFiles[i2], newFilePathCopy);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error coping excels");
                            Console.WriteLine(ex.Message);
                        }
                    }
                    i2++;
                }*/

                // ~~~~~~ find-replace in some excels
                int i3 = 0;
                while (i3 < filesToCreateTemp)
                {
                    if (i3 % 2 == 0)    // to be restored to <<i3 % 3 == 0>> (or not?)
                    {
                        try
                        {
                            workBook = excel.Workbooks.Open(frm1.listOfExcelFiles[i3], m, false, m, m, m, m, m, m, m, m, m, m, m, m);
                            workSheet = (Worksheet)workBook.ActiveSheet;

                            Range rng = (Range)workSheet.UsedRange;

                            rng.Replace(1000, frm1.textToReplace1, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, m, m, m);
                            rng.Replace(2000, frm1.textToReplace2, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, m, m, m);
                            rng.Replace(3000, frm1.textToReplace1, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, m, m, m);
                            rng.Replace(4000, frm1.textToReplace2, XlLookAt.xlWhole, XlSearchOrder.xlByRows, true, m, m, m);

                            workBook.Save();
                            workBook.Close();
                            workBook = null;
                            workSheet = null;
                            if (workSheet != null) Marshal.ReleaseComObject(workSheet);
                            if (workBook != null) Marshal.ReleaseComObject(workBook);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error in find-replace");
                            Console.WriteLine(ex.Message);
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
                Console.WriteLine("backGroundWorker ExcelCreate Error");
                Console.WriteLine(ex.Message);
            }
            tskCreateExcelsRunning = false;
            Console.WriteLine();
            Console.WriteLine("runEverything(): createExcels() -> DONE");
        }
    }
}
