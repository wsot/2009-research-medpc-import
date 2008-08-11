using System;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
    
namespace MedPC_Import
{
    public partial class ThisAddIn
    {
        private Office.CommandBar toolbar;
        private Office.CommandBarButton importButton;
        string xmlFilePath;
        string dataFilePath;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            #region VSTO generated code

            this.Application = (Excel.Application)Microsoft.Office.Tools.Excel.ExcelLocale1033Proxy.Wrap(typeof(Excel.Application), this.Application);

            #endregion
            AddToolbar();

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.toolbar.Delete();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        /**
         * AddToolbar
         * Adds the toolbar and button to Excel
         **/
        private void AddToolbar()
        {
            if (this.toolbar == null)
            {
                Office.CommandBars cmdBar = this.Application.CommandBars;
                this.toolbar = cmdBar.Add("MedPC Import Utility", Office.MsoBarPosition.msoBarTop, false, true);
                this.toolbar.Visible = true;
            }

            try
            {
                importButton = (Office.CommandBarButton)toolbar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, missing);
                importButton.Caption = "Import data";
                importButton.Tag = "MPC_Import";
                importButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(importButton_click);
                importButton.Picture = getImage();

                dataFilePath = "c:\\MED-PC IV\\Data";
                xmlFilePath = "c:\\MED-PC IV\\MPC";
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        
        /**
         * importButton_click
         * Triggered when the 'Import data' button is clicked
         **/
        private void importButton_click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            FileParser theParser;

            //Open file dialog to allow the user to select files
            OpenFileDialog theDialog = new OpenFileDialog();
            if (System.IO.File.Exists(dataFilePath))
                theDialog.InitialDirectory = dataFilePath;
            //theDialog.RestoreDirectory = true;
            theDialog.Multiselect = true; //allow selection of multiple files

            if (theDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    
                    //run the parser on each file selected
                    foreach (string fileName in theDialog.FileNames)
                    {
                        theParser = new FileParser(fileName);
                        theParser.Parse(this.Application, ref xmlFilePath);
                        theParser = null;
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }

        private stdole.IPictureDisp getImage()
        {
            stdole.IPictureDisp tempImage = null;
            try
            {
                System.Drawing.Icon newIcon =
                    Properties.Resources.MedPC;

                ImageList newImageList = new ImageList();
                newImageList.Images.Add(newIcon);
                tempImage = ConvertImage.Convert(newImageList.Images[0]);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return tempImage;
        }
    }
}

sealed public class ConvertImage : System.Windows.Forms.AxHost
{
    private ConvertImage()
        : base(null)
    {
    }
    public static stdole.IPictureDisp Convert
        (System.Drawing.Image image)
    {
        return (stdole.IPictureDisp)System.
            Windows.Forms.AxHost
            .GetIPictureDispFromPicture(image);
    }
}
