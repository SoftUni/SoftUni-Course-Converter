namespace SoftUni_Course_Converter
{
    partial class FormCourseConverter
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
            System.Windows.Forms.Label labelChooseFile;
            System.Windows.Forms.Label labelChooseFolder;
            System.Windows.Forms.Label labelWhatToConvert;
            System.Windows.Forms.Label labelChoosePPTXTemplate;
            System.Windows.Forms.Label labelChooseDOCXTemplate;
            System.Windows.Forms.Label labelFilesToConvert;
            System.Windows.Forms.Label label1;
            System.Windows.Forms.Label labelOutputFolder;
            this.tabControlSrc = new System.Windows.Forms.TabControl();
            this.tabPageSingleDoc = new System.Windows.Forms.TabPage();
            this.buttonChooseFile = new System.Windows.Forms.Button();
            this.textBoxFileToConvert = new System.Windows.Forms.TextBox();
            this.tabPageFolder = new System.Windows.Forms.TabPage();
            this.buttonChooseFolder = new System.Windows.Forms.Button();
            this.textBoxFolderToConvert = new System.Windows.Forms.TextBox();
            this.buttonConvert = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.comboBoxPPTXTemplates = new System.Windows.Forms.ComboBox();
            this.comboBoxDOCXTemplates = new System.Windows.Forms.ComboBox();
            this.listBoxFilesToConvert = new System.Windows.Forms.ListBox();
            this.textBoxLogs = new System.Windows.Forms.TextBox();
            this.textBoxOutputFolder = new System.Windows.Forms.TextBox();
            this.checkBoxSilentConversion = new System.Windows.Forms.CheckBox();
            labelChooseFile = new System.Windows.Forms.Label();
            labelChooseFolder = new System.Windows.Forms.Label();
            labelWhatToConvert = new System.Windows.Forms.Label();
            labelChoosePPTXTemplate = new System.Windows.Forms.Label();
            labelChooseDOCXTemplate = new System.Windows.Forms.Label();
            labelFilesToConvert = new System.Windows.Forms.Label();
            label1 = new System.Windows.Forms.Label();
            labelOutputFolder = new System.Windows.Forms.Label();
            this.tabControlSrc.SuspendLayout();
            this.tabPageSingleDoc.SuspendLayout();
            this.tabPageFolder.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelChooseFile
            // 
            labelChooseFile.AutoSize = true;
            labelChooseFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            labelChooseFile.Location = new System.Drawing.Point(4, 4);
            labelChooseFile.Name = "labelChooseFile";
            labelChooseFile.Size = new System.Drawing.Size(301, 20);
            labelChooseFile.TabIndex = 2;
            labelChooseFile.Text = "Choose file (PPTX / DOCX) to convert:";
            // 
            // labelChooseFolder
            // 
            labelChooseFolder.AutoSize = true;
            labelChooseFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            labelChooseFolder.Location = new System.Drawing.Point(4, 4);
            labelChooseFolder.Name = "labelChooseFolder";
            labelChooseFolder.Size = new System.Drawing.Size(506, 20);
            labelChooseFolder.TabIndex = 5;
            labelChooseFolder.Text = "Choose course folder (holding multiple PPTX / DOCX documents):";
            // 
            // labelWhatToConvert
            // 
            labelWhatToConvert.AutoSize = true;
            labelWhatToConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            labelWhatToConvert.Location = new System.Drawing.Point(7, 9);
            labelWhatToConvert.Name = "labelWhatToConvert";
            labelWhatToConvert.Size = new System.Drawing.Size(676, 20);
            labelWhatToConvert.TabIndex = 1;
            labelWhatToConvert.Text = "Convert PowerPoint presentations and MS Word documents. Choose a document / folde" +
    "r.";
            // 
            // labelChoosePPTXTemplate
            // 
            labelChoosePPTXTemplate.AutoSize = true;
            labelChoosePPTXTemplate.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            labelChoosePPTXTemplate.Location = new System.Drawing.Point(7, 156);
            labelChoosePPTXTemplate.Name = "labelChoosePPTXTemplate";
            labelChoosePPTXTemplate.Size = new System.Drawing.Size(202, 20);
            labelChoosePPTXTemplate.TabIndex = 6;
            labelChoosePPTXTemplate.Text = "Choose a PPTX template:";
            // 
            // labelChooseDOCXTemplate
            // 
            labelChooseDOCXTemplate.AutoSize = true;
            labelChooseDOCXTemplate.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            labelChooseDOCXTemplate.Location = new System.Drawing.Point(7, 197);
            labelChooseDOCXTemplate.Name = "labelChooseDOCXTemplate";
            labelChooseDOCXTemplate.Size = new System.Drawing.Size(208, 20);
            labelChooseDOCXTemplate.TabIndex = 8;
            labelChooseDOCXTemplate.Text = "Choose a DOCX template:";
            // 
            // labelFilesToConvert
            // 
            labelFilesToConvert.AutoSize = true;
            labelFilesToConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            labelFilesToConvert.Location = new System.Drawing.Point(7, 269);
            labelFilesToConvert.Name = "labelFilesToConvert";
            labelFilesToConvert.Size = new System.Drawing.Size(129, 20);
            labelFilesToConvert.TabIndex = 9;
            labelFilesToConvert.Text = "Files to convert:";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            label1.Location = new System.Drawing.Point(7, 436);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(134, 20);
            label1.TabIndex = 11;
            label1.Text = "Conversion logs:";
            // 
            // labelOutputFolder
            // 
            labelOutputFolder.AutoSize = true;
            labelOutputFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            labelOutputFolder.Location = new System.Drawing.Point(7, 237);
            labelOutputFolder.Name = "labelOutputFolder";
            labelOutputFolder.Size = new System.Drawing.Size(179, 20);
            labelOutputFolder.TabIndex = 14;
            labelOutputFolder.Text = "Output (results) folder:";
            // 
            // tabControlSrc
            // 
            this.tabControlSrc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControlSrc.Controls.Add(this.tabPageSingleDoc);
            this.tabControlSrc.Controls.Add(this.tabPageFolder);
            this.tabControlSrc.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabControlSrc.Location = new System.Drawing.Point(11, 39);
            this.tabControlSrc.Name = "tabControlSrc";
            this.tabControlSrc.SelectedIndex = 0;
            this.tabControlSrc.Size = new System.Drawing.Size(1326, 106);
            this.tabControlSrc.TabIndex = 0;
            // 
            // tabPageSingleDoc
            // 
            this.tabPageSingleDoc.BackColor = System.Drawing.SystemColors.Control;
            this.tabPageSingleDoc.Controls.Add(labelChooseFile);
            this.tabPageSingleDoc.Controls.Add(this.buttonChooseFile);
            this.tabPageSingleDoc.Controls.Add(this.textBoxFileToConvert);
            this.tabPageSingleDoc.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabPageSingleDoc.Location = new System.Drawing.Point(4, 29);
            this.tabPageSingleDoc.Name = "tabPageSingleDoc";
            this.tabPageSingleDoc.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSingleDoc.Size = new System.Drawing.Size(1318, 73);
            this.tabPageSingleDoc.TabIndex = 0;
            this.tabPageSingleDoc.Text = "Single Document";
            // 
            // buttonChooseFile
            // 
            this.buttonChooseFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonChooseFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonChooseFile.Location = new System.Drawing.Point(1163, 30);
            this.buttonChooseFile.Name = "buttonChooseFile";
            this.buttonChooseFile.Size = new System.Drawing.Size(143, 32);
            this.buttonChooseFile.TabIndex = 1;
            this.buttonChooseFile.Text = "Choose File";
            this.buttonChooseFile.UseVisualStyleBackColor = true;
            this.buttonChooseFile.Click += new System.EventHandler(this.buttonChooseFile_Click);
            // 
            // textBoxFileToConvert
            // 
            this.textBoxFileToConvert.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxFileToConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxFileToConvert.Location = new System.Drawing.Point(8, 32);
            this.textBoxFileToConvert.Name = "textBoxFileToConvert";
            this.textBoxFileToConvert.ReadOnly = true;
            this.textBoxFileToConvert.Size = new System.Drawing.Size(1144, 27);
            this.textBoxFileToConvert.TabIndex = 0;
            // 
            // tabPageFolder
            // 
            this.tabPageFolder.BackColor = System.Drawing.SystemColors.Control;
            this.tabPageFolder.Controls.Add(labelChooseFolder);
            this.tabPageFolder.Controls.Add(this.buttonChooseFolder);
            this.tabPageFolder.Controls.Add(this.textBoxFolderToConvert);
            this.tabPageFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabPageFolder.Location = new System.Drawing.Point(4, 29);
            this.tabPageFolder.Name = "tabPageFolder";
            this.tabPageFolder.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageFolder.Size = new System.Drawing.Size(1318, 73);
            this.tabPageFolder.TabIndex = 1;
            this.tabPageFolder.Text = "Course Folder";
            // 
            // buttonChooseFolder
            // 
            this.buttonChooseFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonChooseFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonChooseFolder.Location = new System.Drawing.Point(1164, 30);
            this.buttonChooseFolder.Name = "buttonChooseFolder";
            this.buttonChooseFolder.Size = new System.Drawing.Size(143, 32);
            this.buttonChooseFolder.TabIndex = 4;
            this.buttonChooseFolder.Text = "Choose Folder";
            this.buttonChooseFolder.UseVisualStyleBackColor = true;
            this.buttonChooseFolder.Click += new System.EventHandler(this.buttonChooseFolder_Click);
            // 
            // textBoxFolderToConvert
            // 
            this.textBoxFolderToConvert.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxFolderToConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxFolderToConvert.Location = new System.Drawing.Point(8, 32);
            this.textBoxFolderToConvert.Name = "textBoxFolderToConvert";
            this.textBoxFolderToConvert.ReadOnly = true;
            this.textBoxFolderToConvert.Size = new System.Drawing.Size(1144, 27);
            this.textBoxFolderToConvert.TabIndex = 3;
            // 
            // buttonConvert
            // 
            this.buttonConvert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonConvert.Location = new System.Drawing.Point(1197, 156);
            this.buttonConvert.Name = "buttonConvert";
            this.buttonConvert.Size = new System.Drawing.Size(136, 101);
            this.buttonConvert.TabIndex = 3;
            this.buttonConvert.Text = "Convert";
            this.buttonConvert.UseVisualStyleBackColor = true;
            this.buttonConvert.Click += new System.EventHandler(this.buttonConvert_Click);
            // 
            // comboBoxPPTXTemplates
            // 
            this.comboBoxPPTXTemplates.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxPPTXTemplates.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxPPTXTemplates.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBoxPPTXTemplates.FormattingEnabled = true;
            this.comboBoxPPTXTemplates.Location = new System.Drawing.Point(221, 153);
            this.comboBoxPPTXTemplates.Name = "comboBoxPPTXTemplates";
            this.comboBoxPPTXTemplates.Size = new System.Drawing.Size(960, 28);
            this.comboBoxPPTXTemplates.TabIndex = 4;
            // 
            // comboBoxDOCXTemplate
            // 
            this.comboBoxDOCXTemplates.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxDOCXTemplates.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxDOCXTemplates.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBoxDOCXTemplates.FormattingEnabled = true;
            this.comboBoxDOCXTemplates.Location = new System.Drawing.Point(221, 194);
            this.comboBoxDOCXTemplates.Name = "comboBoxDOCXTemplate";
            this.comboBoxDOCXTemplates.Size = new System.Drawing.Size(960, 28);
            this.comboBoxDOCXTemplates.TabIndex = 7;
            // 
            // listBoxFilesToConvert
            // 
            this.listBoxFilesToConvert.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBoxFilesToConvert.BackColor = System.Drawing.SystemColors.Control;
            this.listBoxFilesToConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.listBoxFilesToConvert.FormattingEnabled = true;
            this.listBoxFilesToConvert.ItemHeight = 18;
            this.listBoxFilesToConvert.Location = new System.Drawing.Point(11, 297);
            this.listBoxFilesToConvert.Name = "listBoxFilesToConvert";
            this.listBoxFilesToConvert.Size = new System.Drawing.Size(1322, 130);
            this.listBoxFilesToConvert.TabIndex = 10;
            // 
            // textBoxLogs
            // 
            this.textBoxLogs.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxLogs.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxLogs.Location = new System.Drawing.Point(11, 464);
            this.textBoxLogs.MaxLength = 500000;
            this.textBoxLogs.Multiline = true;
            this.textBoxLogs.Name = "textBoxLogs";
            this.textBoxLogs.ReadOnly = true;
            this.textBoxLogs.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxLogs.Size = new System.Drawing.Size(1322, 219);
            this.textBoxLogs.TabIndex = 12;
            // 
            // textBoxOutputFolder
            // 
            this.textBoxOutputFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxOutputFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxOutputFolder.Location = new System.Drawing.Point(221, 234);
            this.textBoxOutputFolder.Name = "textBoxOutputFolder";
            this.textBoxOutputFolder.Size = new System.Drawing.Size(960, 27);
            this.textBoxOutputFolder.TabIndex = 3;
            // 
            // checkBoxSilentConversion
            // 
            this.checkBoxSilentConversion.AutoSize = true;
            this.checkBoxSilentConversion.Checked = true;
            this.checkBoxSilentConversion.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSilentConversion.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBoxSilentConversion.Location = new System.Drawing.Point(769, 267);
            this.checkBoxSilentConversion.Name = "checkBoxSilentConversion";
            this.checkBoxSilentConversion.Size = new System.Drawing.Size(420, 24);
            this.checkBoxSilentConversion.TabIndex = 15;
            this.checkBoxSilentConversion.Text = "Silent conversion (hide UI for Word and PowerPoint)";
            this.checkBoxSilentConversion.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.checkBoxSilentConversion.UseVisualStyleBackColor = true;
            // 
            // FormCourseConverter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(1347, 695);
            this.Controls.Add(this.checkBoxSilentConversion);
            this.Controls.Add(this.textBoxOutputFolder);
            this.Controls.Add(labelOutputFolder);
            this.Controls.Add(this.textBoxLogs);
            this.Controls.Add(label1);
            this.Controls.Add(this.listBoxFilesToConvert);
            this.Controls.Add(labelFilesToConvert);
            this.Controls.Add(this.comboBoxDOCXTemplates);
            this.Controls.Add(labelChooseDOCXTemplate);
            this.Controls.Add(this.comboBoxPPTXTemplates);
            this.Controls.Add(labelChoosePPTXTemplate);
            this.Controls.Add(this.buttonConvert);
            this.Controls.Add(labelWhatToConvert);
            this.Controls.Add(this.tabControlSrc);
            this.Name = "FormCourseConverter";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SoftUni Course Converter";
            this.Load += new System.EventHandler(this.FormCourseConverter_Load);
            this.tabControlSrc.ResumeLayout(false);
            this.tabPageSingleDoc.ResumeLayout(false);
            this.tabPageSingleDoc.PerformLayout();
            this.tabPageFolder.ResumeLayout(false);
            this.tabPageFolder.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabControlSrc;
        private System.Windows.Forms.TabPage tabPageSingleDoc;
        private System.Windows.Forms.TabPage tabPageFolder;
        private System.Windows.Forms.Button buttonChooseFile;
        private System.Windows.Forms.TextBox textBoxFileToConvert;
        private System.Windows.Forms.Button buttonChooseFolder;
        private System.Windows.Forms.TextBox textBoxFolderToConvert;
        private System.Windows.Forms.Button buttonConvert;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.ComboBox comboBoxPPTXTemplates;
        private System.Windows.Forms.ComboBox comboBoxDOCXTemplates;
        private System.Windows.Forms.ListBox listBoxFilesToConvert;
        private System.Windows.Forms.TextBox textBoxLogs;
        private System.Windows.Forms.TextBox textBoxOutputFolder;
        private System.Windows.Forms.CheckBox checkBoxSilentConversion;
    }
}

