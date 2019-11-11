using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SoftUni_Course_Converter
{
    public partial class FormCourseConverter : Form
    {
        private readonly string docTemplateDir = 
            Directory.GetCurrentDirectory() + @"\..\..\..\Document-Templates";

        public FormCourseConverter()
        {
            InitializeComponent();
        }

        private void FormCourseConverter_Load(object sender, EventArgs e)
        {
            this.openFileDialog.InitialDirectory = Directory.GetCurrentDirectory();
            this.openFileDialog.Filter =
                "PowerPoint Presentations (*.pptx)|*.pptx" + "|" +
                "MS Word Documents (*.docx)|*.docx";

            this.folderBrowserDialog.SelectedPath = Directory.GetCurrentDirectory();

            this.comboBoxPPTXTemplates.Items.AddRange(FindDocTemplates("*.pptx"));
            this.comboBoxPPTXTemplates.SelectedIndex = this.comboBoxPPTXTemplates.Items.Count - 1;

            this.comboBoxDOCXTemplate.Items.AddRange(FindDocTemplates("*.docx"));
            this.comboBoxDOCXTemplate.SelectedIndex = this.comboBoxDOCXTemplate.Items.Count - 1;

            this.textBoxOutputFolder.Text = Path.GetFullPath(
                Directory.GetCurrentDirectory() + @"\..\..\..\output");

            var consoleOutput = new TexBoxWriter(this.textBoxLogs);
            Console.SetOut(consoleOutput);
        }

        private string[] FindDocTemplates(string fileListPattern)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(docTemplateDir);
            FileInfo[] files = dirInfo.GetFiles(fileListPattern);
            string[] templateFileNames = files.Select(f => f.Name).ToArray();
            return templateFileNames;
        }

        private void buttonChooseFile_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxFileToConvert.Text = this.openFileDialog.FileName;
                this.listBoxFilesToConvert.Items.Clear();
                string fullFileName = this.openFileDialog.FileName;
                this.listBoxFilesToConvert.Items.Clear();
                this.listBoxFilesToConvert.Items.Add(fullFileName);
                this.textBoxFolderToConvert.Text = Path.GetDirectoryName(fullFileName);
            }
        }

        private void buttonChooseFolder_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxFolderToConvert.Text = this.folderBrowserDialog.SelectedPath;
                this.listBoxFilesToConvert.Items.Clear();
                string[] filesToConvert = Directory.GetFiles(
                    this.folderBrowserDialog.SelectedPath, "*.*", SearchOption.AllDirectories)
                    .Where(f => f.EndsWith(".pptx") || f.EndsWith(".docx")).ToArray();
                this.listBoxFilesToConvert.Items.Clear();
                this.listBoxFilesToConvert.Items.AddRange(filesToConvert);
            }
        }

        private void buttonConvert_Click(object sender, EventArgs e)
        {
            int filesCount = this.listBoxFilesToConvert.Items.Count;
            if (filesCount == 0)
            {
                Console.WriteLine("Error: no file is selected for conversion.");
                return;
            }

            for (int fileNum = 0; fileNum < filesCount; fileNum++)
            {
                string inputFileName = (string)this.listBoxFilesToConvert.Items[fileNum];
                FileInfo inputFileInfo = new FileInfo(inputFileName);
                Console.WriteLine($"Converting file {inputFileInfo.Name} ({fileNum + 1} of {filesCount})...");
                string inputBaseFolder = this.textBoxFolderToConvert.Text;
                string outputBaseFolder = this.textBoxOutputFolder.Text;
                string outputFileName = inputFileName.Replace(inputBaseFolder, outputBaseFolder);
                string outputFullFolder = new FileInfo(outputFileName).DirectoryName;
                Directory.CreateDirectory(outputFullFolder);
                if (inputFileInfo.Extension.ToLower() == ".pptx")
                {
                    SoftUniPowerPointConverter.ConvertAndFixPresentation(
                        pptSourceFileName: inputFileName,
                        pptDestFileName: outputFileName,
                        pptTemplateFileName: (string)this.comboBoxPPTXTemplates.SelectedItem,
                        appWindowVisible: !this.checkBoxSilentConversion.Checked);
                }
                else if (inputFileInfo.Extension.ToLower() == ".docx")
                {
                    SoftUniMSWordConverter.ConvertAndFixDocument(
                        docSourceFileName: inputFileName,
                        docDestFileName: outputFileName,
                        docTemplateFileName: (string)this.comboBoxPPTXTemplates.SelectedItem,
                        appWindowVisible: !this.checkBoxSilentConversion.Checked);
                }
                else
                {
                    Console.WriteLine($"Unknown file type: {inputFileInfo.Extension}. Conversion skipped.");
                }
            }
        }
    }
}
