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
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WindowsApplication1;
using OfficeUsersSimulation_C;
using System.IO;
using System.Threading;

namespace WindowsApplication1
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                // Command line given, display console
                AllocConsole();
                ConsoleMain(args);
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new OfficeUsersSimulation_C.Form1());
            }
        }
        // ---------------------------------------------------------------------------------------

        // ---------------- Get arguments - check if any missing - save variables ----------------
        private static void ConsoleMain(string[] args)
        {
            cmdVersion cmVer1 = new cmdVersion();
            Form1 frm1 = cmVer1.getForm1();

            for (int arIndx = 0; arIndx < args.Length; arIndx++)
            {
                switch (args[arIndx])
                {
                    /*case "gui":
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);
                        Application.Run(new OfficeUsersSimulation_C.Form1());
                        break; */
                    case "help":
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~ Available arguments ~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                        Console.WriteLine();
                        //Console.WriteLine("gui : start the Graphical Interface");
                        Console.WriteLine("nof + <integer[4-1000]> : number of files to create");
                        Console.WriteLine("emtfls : create empty word and/or excel files");
                        Console.WriteLine("crtfls : create word and/or excel files with text inside");
                        Console.WriteLine("usewd : using word files");
                        Console.WriteLine("usexl : using excel files");
                        Console.WriteLine("vrfrm : vary formatting in files");
                        Console.WriteLine("delaysc + <integer[0-3600]> : delay in seconds during documents creation");
                        Console.WriteLine("autodel : auto delete created files (does NOT delete empty files created)");
                        Console.WriteLine("wrkdir + <path> : working directory");
                        //Console.WriteLine("smwd + <path.docx> : sample word file path");
                        //Console.WriteLine("smxl + <path.xlsx> : sample exel file path");
                        Console.WriteLine();
                        Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine("Press any key to exit...");
                        Console.ReadLine();
                        Environment.Exit(0);        // Exiting
                        Application.Exit();
                        break;
                    case "nof":
                        frm1.filesToCreate = Convert.ToInt32(args[arIndx + 1]);
                        break;
                    case "emtfls":
                        cmVer1.createEmptyFilesCmd = true;
                        break;
                    case "crtfls":
                        cmVer1.createFilesCmd = true;
                        break;
                    case "usewd":
                        frm1.usingWordFiles = true;
                        break;
                    case "usexl":
                        frm1.usingExcelFiles = true;
                        break;
                    case "vrfrm":
                        frm1.varyFormatting = true;
                        break;
                    case "delaysc":
                        if (Convert.ToInt32(args[arIndx + 1]) == 0)
                        {
                            frm1.addDelay = false;
                            frm1.addDelaySec = Convert.ToInt32(args[arIndx + 1]);
                        }
                        else
                        {
                            frm1.addDelay = true;
                            frm1.addDelaySec = Convert.ToInt32(args[arIndx + 1]);
                        }
                        break;
                    case "autodel":
                        frm1.autoDelete = true;
                        break;
                    case "wrkdir":
                        frm1.workingDirectory = Path.GetFullPath(@args[arIndx + 1]);
                        frm1.workingDirectory += "\\";
                        break;
                    /*case "smwd":
                        frm1.sampleWordPath = Path.GetFullPath(@args[arIndx + 1]);
                        break;
                    case "smxl":
                        frm1.sampleExcelPath = Path.GetFullPath(@args[arIndx + 1]);
                        break;*/
                    default:
                        break;
                }
            }
            // checking input data validity
            cmVer1.checkDataValidity();

            // save all settings
            cmVer1.saveAllParameters();

            /*Console.WriteLine();        // for debug!
            cmVer1.printData();
            Console.WriteLine();
            */
            Console.WriteLine();

            if (cmVer1.argumentsOk)
            {
                cmVer1.runEverything();
                cmVer1.waitAllToFinish();
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("printData() result:");
                cmVer1.printData();                 // for debug!

                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("~~~~~ Error(s) encountered! Exiting program....");
                Console.WriteLine();
                Console.WriteLine();
                Console.ReadLine();
                Environment.Exit(0);        // Exiting
                Application.Exit();
            }

            /*
            cmVer1.printData();                 // for debug!
            Console.WriteLine();
            Console.WriteLine("All args using foreach:");
            foreach (string s in args)          // for debug!
            {
                Console.WriteLine(s);
            }
            */
            Console.WriteLine();
            //Console.WriteLine("End of Program.cs"); // for debug!
            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();
            Environment.Exit(0);        // Exiting
            Application.Exit();
        }
        // ---------------------------------------------------------------------------------------

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool AllocConsole();
    }
}


/*              // original code of Program.cs ↓↓
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeUsersSimulation_C
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
*/
