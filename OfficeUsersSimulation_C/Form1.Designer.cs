namespace OfficeUsersSimulation_C
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.buttonRun = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.createEmpyFilesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.filesToCreateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.createEmptyFilesHelpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.workingDirectoryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.wordExcelFilesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.varyFormattingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.delayInEditingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.autoDeleteCreatedFilesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.trackBar1 = new System.Windows.Forms.TrackBar();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.numericUpDownDelayInEditing = new System.Windows.Forms.NumericUpDown();
            this.checkBoxVaryFormating = new System.Windows.Forms.CheckBox();
            this.checkBoxAutoDelete = new System.Windows.Forms.CheckBox();
            this.checkBoxDelayInEditing = new System.Windows.Forms.CheckBox();
            this.checkBoxUsingExcel = new System.Windows.Forms.CheckBox();
            this.checkBoxUsingWord = new System.Windows.Forms.CheckBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.textBoxWorkingDir = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonWorkingDir = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.numericUpDownFilesToCreate = new System.Windows.Forms.NumericUpDown();
            this.textBoxRuntime1 = new System.Windows.Forms.TextBox();
            this.backgroundWorkerWordCreate = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorkerExcelCreate = new System.ComponentModel.BackgroundWorker();
            this.textBoxRuntime2 = new System.Windows.Forms.TextBox();
            this.backgroundWorkerEmptyFiles = new System.ComponentModel.BackgroundWorker();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownDelayInEditing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownFilesToCreate)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonRun
            // 
            this.buttonRun.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonRun.ForeColor = System.Drawing.Color.ForestGreen;
            this.buttonRun.Location = new System.Drawing.Point(12, 245);
            this.buttonRun.Name = "buttonRun";
            this.buttonRun.Size = new System.Drawing.Size(131, 36);
            this.buttonRun.TabIndex = 0;
            this.buttonRun.Text = "R U N";
            this.buttonRun.UseVisualStyleBackColor = true;
            this.buttonRun.Click += new System.EventHandler(this.button1_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem,
            this.createEmpyFilesToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(593, 24);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(52, 20);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // createEmpyFilesToolStripMenuItem
            // 
            this.createEmpyFilesToolStripMenuItem.Enabled = false;
            this.createEmpyFilesToolStripMenuItem.Name = "createEmpyFilesToolStripMenuItem";
            this.createEmpyFilesToolStripMenuItem.Size = new System.Drawing.Size(110, 20);
            this.createEmpyFilesToolStripMenuItem.Text = "Create empy files";
            this.createEmpyFilesToolStripMenuItem.Click += new System.EventHandler(this.createEmpyFilesToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.filesToCreateToolStripMenuItem,
            this.toolStripSeparator1,
            this.createEmptyFilesHelpToolStripMenuItem,
            this.toolStripSeparator3,
            this.workingDirectoryToolStripMenuItem,
            this.toolStripSeparator2,
            this.optionsToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // filesToCreateToolStripMenuItem
            // 
            this.filesToCreateToolStripMenuItem.Name = "filesToCreateToolStripMenuItem";
            this.filesToCreateToolStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.filesToCreateToolStripMenuItem.Text = "Files to create";
            this.filesToCreateToolStripMenuItem.Click += new System.EventHandler(this.filesToCreateToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(166, 6);
            // 
            // createEmptyFilesHelpToolStripMenuItem
            // 
            this.createEmptyFilesHelpToolStripMenuItem.Name = "createEmptyFilesHelpToolStripMenuItem";
            this.createEmptyFilesHelpToolStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.createEmptyFilesHelpToolStripMenuItem.Text = "Create empty files";
            this.createEmptyFilesHelpToolStripMenuItem.Click += new System.EventHandler(this.createEmptyFilesToolStripMenuItem_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(166, 6);
            // 
            // workingDirectoryToolStripMenuItem
            // 
            this.workingDirectoryToolStripMenuItem.Name = "workingDirectoryToolStripMenuItem";
            this.workingDirectoryToolStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.workingDirectoryToolStripMenuItem.Text = "Working directory";
            this.workingDirectoryToolStripMenuItem.Click += new System.EventHandler(this.workingDirectoryToolStripMenuItem_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(166, 6);
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.wordExcelFilesToolStripMenuItem,
            this.varyFormattingToolStripMenuItem,
            this.delayInEditingToolStripMenuItem,
            this.autoDeleteCreatedFilesToolStripMenuItem});
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.optionsToolStripMenuItem.Text = "Options";
            // 
            // wordExcelFilesToolStripMenuItem
            // 
            this.wordExcelFilesToolStripMenuItem.Name = "wordExcelFilesToolStripMenuItem";
            this.wordExcelFilesToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.wordExcelFilesToolStripMenuItem.Text = "Word / Excel files";
            this.wordExcelFilesToolStripMenuItem.Click += new System.EventHandler(this.wordExcelFilesToolStripMenuItem_Click_1);
            // 
            // varyFormattingToolStripMenuItem
            // 
            this.varyFormattingToolStripMenuItem.Name = "varyFormattingToolStripMenuItem";
            this.varyFormattingToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.varyFormattingToolStripMenuItem.Text = "Vary formatting";
            this.varyFormattingToolStripMenuItem.Click += new System.EventHandler(this.varyFormattingToolStripMenuItem_Click);
            // 
            // delayInEditingToolStripMenuItem
            // 
            this.delayInEditingToolStripMenuItem.Name = "delayInEditingToolStripMenuItem";
            this.delayInEditingToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.delayInEditingToolStripMenuItem.Text = "Delay in editing";
            this.delayInEditingToolStripMenuItem.Click += new System.EventHandler(this.delayInEditingToolStripMenuItem_Click);
            // 
            // autoDeleteCreatedFilesToolStripMenuItem
            // 
            this.autoDeleteCreatedFilesToolStripMenuItem.Name = "autoDeleteCreatedFilesToolStripMenuItem";
            this.autoDeleteCreatedFilesToolStripMenuItem.Size = new System.Drawing.Size(201, 22);
            this.autoDeleteCreatedFilesToolStripMenuItem.Text = "Auto delete created files";
            this.autoDeleteCreatedFilesToolStripMenuItem.Click += new System.EventHandler(this.autoDeleteCreatedFilesToolStripMenuItem_Click_1);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(149, 245);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(432, 36);
            this.progressBar1.Step = 1;
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar1.TabIndex = 0;
            // 
            // trackBar1
            // 
            this.trackBar1.LargeChange = 10;
            this.trackBar1.Location = new System.Drawing.Point(215, 28);
            this.trackBar1.Maximum = 1000;
            this.trackBar1.Minimum = 4;
            this.trackBar1.Name = "trackBar1";
            this.trackBar1.Size = new System.Drawing.Size(366, 45);
            this.trackBar1.TabIndex = 10;
            this.trackBar1.Value = 4;
            this.trackBar1.Scroll += new System.EventHandler(this.trackBar1_Scroll);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(8, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 20);
            this.label1.TabIndex = 5;
            this.label1.Text = "Files to create:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.numericUpDownDelayInEditing);
            this.groupBox2.Controls.Add(this.checkBoxVaryFormating);
            this.groupBox2.Controls.Add(this.checkBoxAutoDelete);
            this.groupBox2.Controls.Add(this.checkBoxDelayInEditing);
            this.groupBox2.Controls.Add(this.checkBoxUsingExcel);
            this.groupBox2.Controls.Add(this.checkBoxUsingWord);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(149, 171);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(432, 68);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Options";
            // 
            // numericUpDownDelayInEditing
            // 
            this.numericUpDownDelayInEditing.Location = new System.Drawing.Point(341, 16);
            this.numericUpDownDelayInEditing.Maximum = new decimal(new int[] {
            3600,
            0,
            0,
            0});
            this.numericUpDownDelayInEditing.Name = "numericUpDownDelayInEditing";
            this.numericUpDownDelayInEditing.Size = new System.Drawing.Size(71, 21);
            this.numericUpDownDelayInEditing.TabIndex = 7;
            this.numericUpDownDelayInEditing.Tag = "minutes";
            this.numericUpDownDelayInEditing.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.numericUpDownDelayInEditing.ValueChanged += new System.EventHandler(this.numericUpDown1_ValueChanged);
            // 
            // checkBoxVaryFormating
            // 
            this.checkBoxVaryFormating.AutoSize = true;
            this.checkBoxVaryFormating.Location = new System.Drawing.Point(99, 16);
            this.checkBoxVaryFormating.Name = "checkBoxVaryFormating";
            this.checkBoxVaryFormating.Size = new System.Drawing.Size(107, 19);
            this.checkBoxVaryFormating.TabIndex = 6;
            this.checkBoxVaryFormating.Text = "Vary formatting";
            this.checkBoxVaryFormating.UseVisualStyleBackColor = true;
            this.checkBoxVaryFormating.CheckedChanged += new System.EventHandler(this.checkBox6_CheckedChanged);
            // 
            // checkBoxAutoDelete
            // 
            this.checkBoxAutoDelete.AutoSize = true;
            this.checkBoxAutoDelete.Location = new System.Drawing.Point(225, 43);
            this.checkBoxAutoDelete.Name = "checkBoxAutoDelete";
            this.checkBoxAutoDelete.Size = new System.Drawing.Size(156, 19);
            this.checkBoxAutoDelete.TabIndex = 5;
            this.checkBoxAutoDelete.Text = "Auto delete created files";
            this.checkBoxAutoDelete.UseVisualStyleBackColor = true;
            this.checkBoxAutoDelete.CheckedChanged += new System.EventHandler(this.checkBox5_CheckedChanged);
            // 
            // checkBoxDelayInEditing
            // 
            this.checkBoxDelayInEditing.AutoSize = true;
            this.checkBoxDelayInEditing.Location = new System.Drawing.Point(225, 16);
            this.checkBoxDelayInEditing.Name = "checkBoxDelayInEditing";
            this.checkBoxDelayInEditing.Size = new System.Drawing.Size(110, 19);
            this.checkBoxDelayInEditing.TabIndex = 2;
            this.checkBoxDelayInEditing.Text = "Delay in editing";
            this.checkBoxDelayInEditing.UseVisualStyleBackColor = true;
            this.checkBoxDelayInEditing.CheckedChanged += new System.EventHandler(this.checkBox3_CheckedChanged);
            // 
            // checkBoxUsingExcel
            // 
            this.checkBoxUsingExcel.AutoSize = true;
            this.checkBoxUsingExcel.Location = new System.Drawing.Point(5, 43);
            this.checkBoxUsingExcel.Name = "checkBoxUsingExcel";
            this.checkBoxUsingExcel.Size = new System.Drawing.Size(81, 19);
            this.checkBoxUsingExcel.TabIndex = 1;
            this.checkBoxUsingExcel.Text = "Excel files";
            this.checkBoxUsingExcel.UseVisualStyleBackColor = true;
            this.checkBoxUsingExcel.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // checkBoxUsingWord
            // 
            this.checkBoxUsingWord.AutoSize = true;
            this.checkBoxUsingWord.Location = new System.Drawing.Point(6, 16);
            this.checkBoxUsingWord.Name = "checkBoxUsingWord";
            this.checkBoxUsingWord.Size = new System.Drawing.Size(80, 19);
            this.checkBoxUsingWord.TabIndex = 0;
            this.checkBoxUsingWord.Text = "Word files";
            this.checkBoxUsingWord.UseVisualStyleBackColor = true;
            this.checkBoxUsingWord.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // textBoxWorkingDir
            // 
            this.textBoxWorkingDir.Location = new System.Drawing.Point(164, 79);
            this.textBoxWorkingDir.Name = "textBoxWorkingDir";
            this.textBoxWorkingDir.Size = new System.Drawing.Size(349, 20);
            this.textBoxWorkingDir.TabIndex = 14;
            this.textBoxWorkingDir.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            this.textBoxWorkingDir.Leave += new System.EventHandler(this.textBox2_Leave);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(139, 20);
            this.label2.TabIndex = 15;
            this.label2.Text = "Working directory :";
            // 
            // buttonWorkingDir
            // 
            this.buttonWorkingDir.Location = new System.Drawing.Point(519, 79);
            this.buttonWorkingDir.Name = "buttonWorkingDir";
            this.buttonWorkingDir.Size = new System.Drawing.Size(62, 20);
            this.buttonWorkingDir.TabIndex = 16;
            this.buttonWorkingDir.Text = "Browse...";
            this.buttonWorkingDir.UseVisualStyleBackColor = true;
            this.buttonWorkingDir.Click += new System.EventHandler(this.button2_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "docx";
            this.openFileDialog1.Filter = "Documents (*.docx)|*.docx";
            this.openFileDialog1.Title = "Select a Word file";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.DefaultExt = "xlsx";
            this.openFileDialog2.Filter = "Excel (*.xlsx)|*.xlsx";
            this.openFileDialog2.Title = "Select an Excel file";
            // 
            // numericUpDownFilesToCreate
            // 
            this.numericUpDownFilesToCreate.Location = new System.Drawing.Point(164, 31);
            this.numericUpDownFilesToCreate.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numericUpDownFilesToCreate.Minimum = new decimal(new int[] {
            4,
            0,
            0,
            0});
            this.numericUpDownFilesToCreate.Name = "numericUpDownFilesToCreate";
            this.numericUpDownFilesToCreate.Size = new System.Drawing.Size(45, 20);
            this.numericUpDownFilesToCreate.TabIndex = 23;
            this.numericUpDownFilesToCreate.Value = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numericUpDownFilesToCreate.ValueChanged += new System.EventHandler(this.numericUpDown2_ValueChanged);
            // 
            // textBoxRuntime1
            // 
            this.textBoxRuntime1.Enabled = false;
            this.textBoxRuntime1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxRuntime1.Location = new System.Drawing.Point(12, 57);
            this.textBoxRuntime1.Multiline = true;
            this.textBoxRuntime1.Name = "textBoxRuntime1";
            this.textBoxRuntime1.ReadOnly = true;
            this.textBoxRuntime1.Size = new System.Drawing.Size(569, 60);
            this.textBoxRuntime1.TabIndex = 24;
            this.textBoxRuntime1.Visible = false;
            // 
            // backgroundWorkerWordCreate
            // 
            this.backgroundWorkerWordCreate.WorkerReportsProgress = true;
            this.backgroundWorkerWordCreate.WorkerSupportsCancellation = true;
            this.backgroundWorkerWordCreate.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerWordCreate_DoWork);
            this.backgroundWorkerWordCreate.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerWordCreate_RunWorkerCompleted);
            // 
            // backgroundWorkerExcelCreate
            // 
            this.backgroundWorkerExcelCreate.WorkerReportsProgress = true;
            this.backgroundWorkerExcelCreate.WorkerSupportsCancellation = true;
            this.backgroundWorkerExcelCreate.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerExcelCreate_DoWork);
            this.backgroundWorkerExcelCreate.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerExcelCreate_RunWorkerCompleted);
            // 
            // textBoxRuntime2
            // 
            this.textBoxRuntime2.Enabled = false;
            this.textBoxRuntime2.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxRuntime2.Location = new System.Drawing.Point(12, 114);
            this.textBoxRuntime2.Multiline = true;
            this.textBoxRuntime2.Name = "textBoxRuntime2";
            this.textBoxRuntime2.ReadOnly = true;
            this.textBoxRuntime2.Size = new System.Drawing.Size(569, 125);
            this.textBoxRuntime2.TabIndex = 25;
            this.textBoxRuntime2.Visible = false;
            // 
            // backgroundWorkerEmptyFiles
            // 
            this.backgroundWorkerEmptyFiles.WorkerReportsProgress = true;
            this.backgroundWorkerEmptyFiles.WorkerSupportsCancellation = true;
            this.backgroundWorkerEmptyFiles.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerEmptyFiles_DoWork);
            this.backgroundWorkerEmptyFiles.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerEmptyFiles_RunWorkerCompleted);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(593, 290);
            this.Controls.Add(this.numericUpDownFilesToCreate);
            this.Controls.Add(this.buttonWorkingDir);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxWorkingDir);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.trackBar1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.buttonRun);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.textBoxRuntime1);
            this.Controls.Add(this.textBoxRuntime2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Multiple office users simulator";
            this.MouseHover += new System.EventHandler(this.Form1_MouseHover);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownDelayInEditing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownFilesToCreate)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonRun;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TrackBar trackBar1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox checkBoxUsingExcel;
        private System.Windows.Forms.CheckBox checkBoxUsingWord;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TextBox textBoxWorkingDir;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonWorkingDir;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.CheckBox checkBoxDelayInEditing;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.CheckBox checkBoxAutoDelete;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem filesToCreateToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem wordExcelFilesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem delayInEditingToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem autoDeleteCreatedFilesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem varyFormattingToolStripMenuItem;
        private System.Windows.Forms.CheckBox checkBoxVaryFormating;
        private System.Windows.Forms.NumericUpDown numericUpDownDelayInEditing;
        private System.Windows.Forms.NumericUpDown numericUpDownFilesToCreate;
        private System.Windows.Forms.TextBox textBoxRuntime1;
        public System.ComponentModel.BackgroundWorker backgroundWorkerWordCreate;
        public System.ComponentModel.BackgroundWorker backgroundWorkerExcelCreate;
        private System.Windows.Forms.TextBox textBoxRuntime2;
        private System.Windows.Forms.ToolStripMenuItem createEmpyFilesToolStripMenuItem;
        public System.ComponentModel.BackgroundWorker backgroundWorkerEmptyFiles;
        private System.Windows.Forms.ToolStripMenuItem createEmptyFilesHelpToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripMenuItem workingDirectoryToolStripMenuItem;
    }
}

