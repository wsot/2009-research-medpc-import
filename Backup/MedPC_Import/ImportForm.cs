using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MedPC_Import
{
    public partial class ImportForm : Form
    {
        private string dataFilePath;
        private string xmlFilePath;
        private string prevOutputFilename;
        private string prevOutputPath;

        public ImportForm(string theDataFilePath, string theXmlFilePath)
        {
            InitializeComponent();
            dataFilePath = theDataFilePath;
            xmlFilePath = theXmlFilePath;
        }

        private void importToSameFolder_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked)
            {
                prevOutputPath = importDestination.Text;
                importDestination.Text = "Same as source file";
                importDestination.Enabled = false;
                importToLabel.Enabled = false;
            }
            else
            {
                if (prevOutputPath == null)
                    importDestination.Text = "";
                else
                    importDestination.Text = prevOutputPath;
                importDestination.Enabled = true;
                importToLabel.Enabled = true;
            }
        }

        private void addFileButton_Click(object sender, EventArgs e)
        {
            //Open file dialog to allow the user to select files
            OpenFileDialog theDialog = new OpenFileDialog();
            if (System.IO.File.Exists(dataFilePath))
                theDialog.InitialDirectory = dataFilePath;
            theDialog.Multiselect = true; //allow selection of multiple files

            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    this.fileList.Items.AddRange(theDialog.FileNames);
                }
                catch (Exception ea)
                {
                    MessageBox.Show(ea.Message);
                }
            }
        }

        private void removeFileButton_Click(object sender, EventArgs e)
        {
            while (fileList.SelectedIndices.Count > 0)
            {
                fileList.Items.RemoveAt(fileList.SelectedIndices[0]);
            }
        }

        private void useCustomFilename_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked)
            {
                prevOutputFilename = outputFilenameField.Text;
                outputFilenameField.Text = "Same as input filename";
                outputFilenameField.Enabled = false;
                outputFilenameLabel.Enabled = false;
            }
            else
            {
                if (prevOutputFilename == null)
                    outputFilenameField.Text = "";
                else
                    outputFilenameField.Text = prevOutputFilename;
                outputFilenameField.Enabled = true;
                outputFilenameLabel.Enabled = true;
            }
        }

        private void outputFilenameLabel_Click(object sender, EventArgs e)
        {

        }
    }
}
