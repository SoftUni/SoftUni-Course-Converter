using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
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
                "Presentations and Documents (*.pptx; *.docx)|*.pptx;*.docx";

            this.folderBrowserDialog.SelectedPath = Directory.GetCurrentDirectory();

            this.comboBoxPPTXTemplates.Items.AddRange(FindDocTemplates("*.pptx"));
            this.comboBoxPPTXTemplates.SelectedIndex = this.comboBoxPPTXTemplates.Items.Count - 1;

            this.comboBoxDOCXTemplates.Items.AddRange(FindDocTemplates("*.docx"));
            this.comboBoxDOCXTemplates.SelectedIndex = this.comboBoxDOCXTemplates.Items.Count - 1;

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
                    this.folderBrowserDialog.SelectedPath, "*.*", SearchOption.AllDirectories).ToArray();
                this.listBoxFilesToConvert.Items.Clear();
                this.listBoxFilesToConvert.Items.AddRange(filesToConvert);
            }
        }

        private async void buttonConvert_Click(object sender, EventArgs e)
        {
            this.textBoxLogs.Clear();

            int filesCount = this.listBoxFilesToConvert.Items.Count;
            if (filesCount == 0)
            {
                Console.WriteLine("Error: no files / folders are selected for conversion.");
                return;
            }

            this.buttonConvert.Enabled = false;

            for (int fileNum = 0; fileNum < filesCount; fileNum++)
            {
                string inputFileName = (string)this.listBoxFilesToConvert.Items[fileNum];
                this.listBoxFilesToConvert.SelectedIndex = fileNum;
                FileInfo inputFileInfo = new FileInfo(inputFileName);
                Console.WriteLine($"Converting file {inputFileInfo.Name} ({fileNum + 1} of {filesCount})...");
                bool silentConversion = this.checkBoxSilentConversion.Checked;
                string inputBaseFolder = this.textBoxFolderToConvert.Text;
                string outputBaseFolder = this.textBoxOutputFolder.Text;
                string outputFileName = inputFileName.Replace(inputBaseFolder, outputBaseFolder);
                string outputFullFolder = new FileInfo(outputFileName).DirectoryName;
                Directory.CreateDirectory(outputFullFolder);
                try
                {
                    if (inputFileInfo.Extension.ToLower() == ".pptx")
                    {
                        string templateFileName = Path.GetFullPath(docTemplateDir + @"\" +
                            (string)this.comboBoxPPTXTemplates.SelectedItem);
                        await Task.Run(() =>
                        {
                            SoftUniPowerPointConverter.ConvertAndFixPresentation(
                                pptSourceFileName: inputFileName,
                                pptDestFileName: outputFileName,
                                pptTemplateFileName: templateFileName,
                                appWindowVisible: !silentConversion);
                        });
                    }
                    else if (inputFileInfo.Extension.ToLower() == ".docx")
                    {
                        string templateFileName = Path.GetFullPath(docTemplateDir + @"\" +
                            (string)this.comboBoxDOCXTemplates.SelectedItem);
                        await Task.Run(() =>
                        {
                            SoftUniMSWordConverter.ConvertAndFixDocument(
                                docSourceFileName: inputFileName,
                                docDestFileName: outputFileName,
                                docTemplateFileName: templateFileName,
                                appWindowVisible: !silentConversion);
                        });
                    }
                    else
                    {
                        File.Copy(inputFileName, outputFileName, true);
                        Console.WriteLine($"Unknown file type: {inputFileInfo.Name}. Stored to the output folder.");
                    }
                    Console.WriteLine($"Conversion of file {inputFileInfo.Name} completed.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting file {inputFileInfo.Name}: {ex.Message}");
                    Console.WriteLine(ex.StackTrace);
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                Console.WriteLine();
            }

            // Conversion complated
            Console.WriteLine("All files converted successfully.");
            this.buttonConvert.Enabled = true;
        }
    }
}
