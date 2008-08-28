namespace MedPC_Import
{
    partial class ImportForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ImportForm));
            this.fileList = new System.Windows.Forms.ListBox();
            this.filesToImportLabel = new System.Windows.Forms.Label();
            this.addFileButton = new System.Windows.Forms.Button();
            this.removeFileButton = new System.Windows.Forms.Button();
            this.importToSameFolder = new System.Windows.Forms.CheckBox();
            this.importDestination = new System.Windows.Forms.TextBox();
            this.importToLabel = new System.Windows.Forms.Label();
            this.importTemplateFolder = new System.Windows.Forms.TextBox();
            this.templateFolderLabel = new System.Windows.Forms.Label();
            this.importButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.useCustomFilename = new System.Windows.Forms.CheckBox();
            this.outputFilenameLabel = new System.Windows.Forms.Label();
            this.outputFilenameField = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // fileList
            // 
            this.fileList.FormattingEnabled = true;
            this.fileList.Location = new System.Drawing.Point(12, 33);
            this.fileList.Name = "fileList";
            this.fileList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.fileList.Size = new System.Drawing.Size(488, 264);
            this.fileList.TabIndex = 0;
            // 
            // filesToImportLabel
            // 
            this.filesToImportLabel.AutoSize = true;
            this.filesToImportLabel.Location = new System.Drawing.Point(13, 14);
            this.filesToImportLabel.Name = "filesToImportLabel";
            this.filesToImportLabel.Size = new System.Drawing.Size(74, 13);
            this.filesToImportLabel.TabIndex = 1;
            this.filesToImportLabel.Text = "Files to import:";
            // 
            // addFileButton
            // 
            this.addFileButton.Location = new System.Drawing.Point(506, 33);
            this.addFileButton.Name = "addFileButton";
            this.addFileButton.Size = new System.Drawing.Size(124, 23);
            this.addFileButton.TabIndex = 2;
            this.addFileButton.Text = "&Add files";
            this.addFileButton.UseVisualStyleBackColor = true;
            this.addFileButton.Click += new System.EventHandler(this.addFileButton_Click);
            // 
            // removeFileButton
            // 
            this.removeFileButton.Location = new System.Drawing.Point(506, 63);
            this.removeFileButton.Name = "removeFileButton";
            this.removeFileButton.Size = new System.Drawing.Size(124, 23);
            this.removeFileButton.TabIndex = 3;
            this.removeFileButton.Text = "&Remove selected";
            this.removeFileButton.UseVisualStyleBackColor = true;
            this.removeFileButton.Click += new System.EventHandler(this.removeFileButton_Click);
            // 
            // importToSameFolder
            // 
            this.importToSameFolder.AutoSize = true;
            this.importToSameFolder.Checked = true;
            this.importToSameFolder.CheckState = System.Windows.Forms.CheckState.Checked;
            this.importToSameFolder.Location = new System.Drawing.Point(16, 304);
            this.importToSameFolder.Name = "importToSameFolder";
            this.importToSameFolder.Size = new System.Drawing.Size(194, 17);
            this.importToSameFolder.TabIndex = 4;
            this.importToSameFolder.Text = "Import to same folder as source files";
            this.importToSameFolder.UseVisualStyleBackColor = true;
            this.importToSameFolder.CheckedChanged += new System.EventHandler(this.importToSameFolder_CheckedChanged);
            // 
            // importDestination
            // 
            this.importDestination.Enabled = false;
            this.importDestination.Location = new System.Drawing.Point(122, 328);
            this.importDestination.Name = "importDestination";
            this.importDestination.Size = new System.Drawing.Size(378, 20);
            this.importDestination.TabIndex = 5;
            this.importDestination.Text = "Same as source file";
            // 
            // importToLabel
            // 
            this.importToLabel.AutoSize = true;
            this.importToLabel.Enabled = false;
            this.importToLabel.Location = new System.Drawing.Point(36, 331);
            this.importToLabel.Name = "importToLabel";
            this.importToLabel.Size = new System.Drawing.Size(80, 13);
            this.importToLabel.TabIndex = 6;
            this.importToLabel.Text = "Import to folder:";
            this.importToLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // importTemplateFolder
            // 
            this.importTemplateFolder.Location = new System.Drawing.Point(122, 412);
            this.importTemplateFolder.Name = "importTemplateFolder";
            this.importTemplateFolder.Size = new System.Drawing.Size(378, 20);
            this.importTemplateFolder.TabIndex = 7;
            // 
            // templateFolderLabel
            // 
            this.templateFolderLabel.AutoSize = true;
            this.templateFolderLabel.Location = new System.Drawing.Point(33, 415);
            this.templateFolderLabel.Name = "templateFolderLabel";
            this.templateFolderLabel.Size = new System.Drawing.Size(83, 13);
            this.templateFolderLabel.TabIndex = 8;
            this.templateFolderLabel.Text = "Template folder:";
            this.templateFolderLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // importButton
            // 
            this.importButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.importButton.Location = new System.Drawing.Point(16, 445);
            this.importButton.Name = "importButton";
            this.importButton.Size = new System.Drawing.Size(75, 23);
            this.importButton.TabIndex = 9;
            this.importButton.Text = "&Import";
            this.importButton.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(97, 445);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 10;
            this.cancelButton.Text = "&Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // useCustomFilename
            // 
            this.useCustomFilename.AutoSize = true;
            this.useCustomFilename.Checked = true;
            this.useCustomFilename.CheckState = System.Windows.Forms.CheckState.Checked;
            this.useCustomFilename.Location = new System.Drawing.Point(16, 356);
            this.useCustomFilename.Name = "useCustomFilename";
            this.useCustomFilename.Size = new System.Drawing.Size(203, 17);
            this.useCustomFilename.TabIndex = 11;
            this.useCustomFilename.Text = "Use input filename for output filename";
            this.useCustomFilename.UseVisualStyleBackColor = true;
            this.useCustomFilename.CheckedChanged += new System.EventHandler(this.useCustomFilename_CheckedChanged);
            // 
            // outputFilenameLabel
            // 
            this.outputFilenameLabel.AutoSize = true;
            this.outputFilenameLabel.Enabled = false;
            this.outputFilenameLabel.Location = new System.Drawing.Point(29, 383);
            this.outputFilenameLabel.Name = "outputFilenameLabel";
            this.outputFilenameLabel.Size = new System.Drawing.Size(87, 13);
            this.outputFilenameLabel.TabIndex = 12;
            this.outputFilenameLabel.Text = "Output file name:";
            this.outputFilenameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.outputFilenameLabel.Click += new System.EventHandler(this.outputFilenameLabel_Click);
            // 
            // outputFilenameField
            // 
            this.outputFilenameField.Enabled = false;
            this.outputFilenameField.Location = new System.Drawing.Point(123, 380);
            this.outputFilenameField.Name = "outputFilenameField";
            this.outputFilenameField.Size = new System.Drawing.Size(377, 20);
            this.outputFilenameField.TabIndex = 13;
            this.outputFilenameField.Text = "Same as input filename";
            // 
            // ImportForm
            // 
            this.AcceptButton = this.importButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(638, 476);
            this.Controls.Add(this.outputFilenameField);
            this.Controls.Add(this.outputFilenameLabel);
            this.Controls.Add(this.useCustomFilename);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.importButton);
            this.Controls.Add(this.templateFolderLabel);
            this.Controls.Add(this.importTemplateFolder);
            this.Controls.Add(this.importToLabel);
            this.Controls.Add(this.importDestination);
            this.Controls.Add(this.importToSameFolder);
            this.Controls.Add(this.removeFileButton);
            this.Controls.Add(this.addFileButton);
            this.Controls.Add(this.filesToImportLabel);
            this.Controls.Add(this.fileList);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ImportForm";
            this.Text = "ImportForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox fileList;
        private System.Windows.Forms.Label filesToImportLabel;
        private System.Windows.Forms.Button addFileButton;
        private System.Windows.Forms.Button removeFileButton;
        private System.Windows.Forms.CheckBox importToSameFolder;
        private System.Windows.Forms.TextBox importDestination;
        private System.Windows.Forms.Label importToLabel;
        private System.Windows.Forms.TextBox importTemplateFolder;
        private System.Windows.Forms.Label templateFolderLabel;
        private System.Windows.Forms.Button importButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.CheckBox useCustomFilename;
        private System.Windows.Forms.Label outputFilenameLabel;
        private System.Windows.Forms.TextBox outputFilenameField;
    }
}