using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace MedPC_Import
{
    class FileParser
    {
        System.IO.StreamReader stream;
        Excel.Workbook wb;
        string msn; //MSN is the name of the program which was run - used to locate the XML description file
        string dataFilename; //the name of the data file taken from the file header; stored to put in each excel output file
        string inputFilename; //the actual filename of the file being parsed; used for location to which file should be saved.
        string line; //the current line in the file being parsed
        string xmlFilePath; //path where the XML files can be found
        int headerCount; //used to track the number of headers found in the current file
        System.Collections.Hashtable xmlVars;
        System.Collections.Hashtable xmlArrays;
        Excel.Application app;

        public FileParser(string theFilename)
        {
            inputFilename = theFilename;
            stream = new System.IO.StreamReader(theFilename);
        }

        public void Parse(Excel.Application theApp, ref string theXmlPath)
        {
            xmlFilePath = theXmlPath;
            app = theApp;
            
            while (!stream.EndOfStream)
            {
                xmlVars = new System.Collections.Hashtable();
                xmlArrays = new System.Collections.Hashtable();

                line = stream.ReadLine();
                string outputFilename;

                parseHeader();
                if (!readXml())
                {
                    MessageBox.Show(String.Concat("Reading file '", inputFilename, "' failed - XML descriptor not found"));
                    break;
                }
                parseVariables();

                xmlVars = null;
                xmlArrays = null;

                if (wb != null)
                {
                    if (headerCount > 1)
                    {
                        outputFilename = String.Concat(inputFilename, '_', Convert.ToString(headerCount), ".xls");
                    }
                    else
                    {
                        outputFilename = String.Concat(inputFilename, ".xls");
                    }
                    try
                    {
                        wb.SaveAs(outputFilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                    catch (Exception e)
                    {
//                        MessageBox.Show(e.Message + "\n" + e.StackTrace);
                    }

                }
            }

            app = null;
            stream.Close();
            theXmlPath = xmlFilePath;
        }

        /***
         * Parses the MedPC header data from an output file
         **/
        private void parseHeader()
        {
            wb = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet); //create new workbook for the data
            Excel.Worksheet wsOverview = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Excel.XlSheetType.xlWorksheet); //create a worksheet for header information
            wsOverview.Name = "Overview";

            int colonPos = 0, rowNum = 1;
            string label, value;

            while (!stream.EndOfStream)
            {
                colonPos = line.IndexOf(':');
                if (colonPos != -1)
                {
                    if (rowNum == 1) headerCount++;

                    label = line.Substring(0, colonPos);
                    value = line.Substring(colonPos + 1, line.Length - colonPos - 1).Trim();
                    if (label.Equals("File"))
                    {
                        dataFilename = value;
                    } else if (rowNum == 1) {
                        wsOverview.get_Range(String.Concat('A', rowNum.ToString()), Type.Missing).Value2 = "File";
                        wsOverview.get_Range(String.Concat('B', rowNum.ToString()), Type.Missing).Value2 = dataFilename;
                        rowNum++;
                    }

                    wsOverview.get_Range(String.Concat('A', rowNum.ToString()), Type.Missing).Value2 = label;
                    wsOverview.get_Range(String.Concat('B', rowNum.ToString()), Type.Missing).Value2 = value;
                    rowNum++;
                    if (label.Equals("MSN"))
                    {
                        msn = value;
                        break;
                    }
                }
                line = stream.ReadLine();
            }
        }

        /**
         * Parse variables and arrays out of the file.
         * Only those variable names matching an entry in xmlVars and array names matching an entry is xmlArrays will be includeded
         */
        private void parseVariables()
        {
            try
            {
                Excel.Worksheet wsVars = null;

                int rowNum = 1;
                string label, value;
                value = "";

                Regex varStartRegex = new Regex(@"^\s*([A-Za-z]):\s*(?:(-?[0-9\.]+)\s*)?$", RegexOptions.IgnoreCase); //regex will match start  of variable or array - the second group only captures for variable, not array

                line = stream.ReadLine();
                while (!stream.EndOfStream && varStartRegex.IsMatch(line)) //match -> a variable name (and possibly value) is present on this line
                {
                    Match theMatch = varStartRegex.Match(line);
                    label = theMatch.Groups[1].Captures[0].Value;

                    if (theMatch.Groups[2].Captures.Count > 0) //if >0 then variable value on same line. Otherwise it is an array.
                    {
                        label = theMatch.Groups[1].Value;
                        value = theMatch.Groups[2].Value;

                        if (xmlVars.ContainsKey(label)) //check variable is in list of vars to output
                        {
                            if (wsVars == null) //if there is no 'variables' worksheet, create it.
                            {
                                wsVars = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Excel.XlSheetType.xlWorksheet);
                                wsVars.Name = "Variables";
                            }
                            //put variable in the worksheet
                            wsVars.get_Range(String.Concat('A', rowNum.ToString()), Type.Missing).Value2 = ((MPCVariable)xmlVars[label]).outputName;
                            wsVars.get_Range(String.Concat('B', rowNum.ToString()), Type.Missing).Value2 = value;
                            wsVars.get_Range(String.Concat('C', rowNum.ToString()), Type.Missing).Value2 = ((MPCVariable)xmlVars[label]).summary;
                            rowNum++;
                        }

                        line = stream.ReadLine();
                    }
                    else
                    {
                        //line contains beginning of an array
                        label = theMatch.Groups[1].Value;
                        if (xmlArrays.ContainsKey(label)) //check if the array is one that should be output
                        {
                            readDataArray(label);
                        }
                        else
                        {
                            //array is not one to be output, so skip to end of array (actually comes back with line = the line of text after the end of the array)
                            line = stream.ReadLine();
                            skipArray();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        /**
         * Read data from the array and output onto a new worksheet specifically for the array
         */
        private void readDataArray(string label)
        {
            int itemNum = 0;
            MPCArray theArray = (MPCArray)xmlArrays[label];
            MPCArrayColumn theCol;

            Excel.Worksheet arrWs = (Excel.Worksheet) wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Excel.XlSheetType.xlWorksheet);
            if (theArray.outputName != null)
            {
                arrWs.Name = theArray.outputName;
            }
            else
            {
                arrWs.Name = String.Concat("Array ", label);
            }

            int columnCount = theArray.columns.Count;
            outputColHeaders(theArray, arrWs);
            int colNum, rowNum;

            //Regex matches a five-column output. If not 5-column, this will break.
            Regex arrayValuesRegex = new Regex(@"^\s*(\d+):\s*(-?[0-9\.]+)\s*(?:(-?[0-9\.]+)\s*)?(?:(-?[0-9\.]+)\s*)?(?:(-?[0-9\.]+)\s*)?(?:(-?[0-9\.]+)\s*)?$", RegexOptions.IgnoreCase);

            line = stream.ReadLine();
            while (!stream.EndOfStream && arrayValuesRegex.IsMatch(line)) //check we are still in an array and not at end of file
            {
                Match theMatch = arrayValuesRegex.Match(line);
                for (int i = 2; i < theMatch.Groups.Count; i++) //output data for each capture (which can be <5 if we are on the last line of the array)
                {
                    colNum = (itemNum % columnCount);
                    rowNum = (itemNum / columnCount) + 2;
                    if (theMatch.Groups[i].Captures.Count > 0)
                    {
                        theCol = (MPCArrayColumn)theArray.columns[colNum];
                        if (theCol.includeInOutput)
                        {
                            if (theArray.outputStyle == "cols")
                            {
                                //output each data 'row' in a vertical column
                                arrWs.get_Range(String.Concat(getColName(rowNum), System.Convert.ToString(theCol.outputColNum)), Type.Missing).Value2 = theMatch.Groups[i].Value;
                            }
                            else
                            {
                                //output each data 'row' in a horizontal row
                                arrWs.get_Range(String.Concat(getColName(theCol.outputColNum), System.Convert.ToString(rowNum)), Type.Missing).Value2 = theMatch.Groups[i].Value;
                            }
                        }
                        itemNum++;
                    }
                }
                line = stream.ReadLine();
            }
        }

        /**
         * Read the XML file associated with the program run. If the XML file is missing, then will ask for where to locate it
         */
        private bool readXml()
        {
            bool fileLoaded = false;

            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
            try
            {
                //check default location C:\MED-PC IV\MPC\<msn>.MPC.xml
                if (System.IO.File.Exists(String.Concat(xmlFilePath, System.IO.Path.DirectorySeparatorChar, msn, ".MPC.xml")))
                {
                    xmlDoc.Load(String.Concat(xmlFilePath, System.IO.Path.DirectorySeparatorChar, msn, ".MPC.xml"));
                    fileLoaded = true;
                }

                //check input file location for <msn>.MPC.xml
                if (System.IO.File.Exists(String.Concat(System.IO.Path.GetDirectoryName(inputFilename), System.IO.Path.DirectorySeparatorChar, msn, ".MPC.xml")))
                {
                    xmlDoc.Load(String.Concat(System.IO.Path.GetDirectoryName(inputFilename), System.IO.Path.DirectorySeparatorChar, msn, ".MPC.xml"));
                    fileLoaded = true;
                }

            }
            catch (System.IO.FileNotFoundException fnfe)
            { }
            catch (System.Xml.XmlException xmle)
            { }
            catch (System.IO.IOException ioe)
            { }

            if (!fileLoaded)
            {
                //didn't exist in default location - ask user where the file is located
                OpenFileDialog theDialog = new OpenFileDialog();
                if (System.IO.File.Exists(xmlFilePath))
                    theDialog.InitialDirectory = xmlFilePath;

                if (theDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        xmlDoc.Load(theDialog.FileName);
                        xmlFilePath = System.IO.Path.GetDirectoryName(theDialog.FileName);
                        fileLoaded = true;
                    }
                    catch (System.IO.FileNotFoundException fnfe)
                    { }
                    catch (System.Xml.XmlException xmle)
                    { }
                    catch (System.IO.IOException ioe)
                    { }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message + "\n" + e.StackTrace);
                        fileLoaded = false;
                    }
                }
            }

            if (fileLoaded)
            {
                try
                {
                    //Process the variables in the file
                    System.Xml.XmlNodeList nodes = xmlDoc.SelectNodes("/program/variable");
                    foreach (System.Xml.XmlNode node in nodes)
                    {
                        //if outputName exists, then the variable should be included
                        if (node.Attributes["outputName"] != null)
                        {
                            MPCVariable theVar = new MPCVariable();
                            theVar.label = node.Attributes["name"].Value;
                            theVar.summary = node.Attributes["summary"].Value;
                            theVar.outputName = node.Attributes["outputName"].Value;
                            xmlVars.Add(theVar.label, theVar);
                        }
                    }

                    nodes = xmlDoc.SelectNodes("program/array");
                    foreach (System.Xml.XmlNode node in nodes)
                    {
                        if (node.Attributes["name"] != null)
                        {
                            bool includeInOutput = false;
                            MPCArray theArray = new MPCArray();
                            int outputColNum = 0;
                            theArray.name = node.Attributes["name"].Value;
                            theArray.summary = node.Attributes["summary"].Value;
                            //if outputName exists, then the array should be included
                            if (node.Attributes["outputName"] != null)
                            {
                                theArray.outputName = node.Attributes["outputName"].Value;
                                if (node.Attributes["outputStyle"] != null)
                                {
                                    //if an output style (cols or rows) is specified, use it. cols = each now data 'row' is a new column; rows = each new data 'row' is a row.
                                    theArray.outputStyle = node.Attributes["outputStyle"].Value;
                                }
                                else
                                {
                                    //by default, output so each data 'line' is on a new row
                                    theArray.outputStyle = "rows";
                                }
                            }
                            theArray.columns = new System.Collections.ArrayList();

                            //create the list of columns in the array
                            foreach (System.Xml.XmlNode columnNode in node.ChildNodes)
                            {
                                MPCArrayColumn theColumn = new MPCArrayColumn();
                                theColumn.name = columnNode.Attributes["name"].Value;
                                if (columnNode.Attributes["outputName"] != null)
                                {
                                    outputColNum++;
                                    theColumn.outputName = columnNode.Attributes["outputName"].Value;
                                    theColumn.outputColNum = outputColNum;
                                    theColumn.includeInOutput = true;
                                    includeInOutput = true;
                                }
                                else
                                {
                                    theColumn.includeInOutput = false;
                                }
                                theColumn.summary = columnNode.Attributes["summary"].Value;
                                theArray.columns.Add(theColumn);
                            }

                            if (includeInOutput)
                            {
                                xmlArrays.Add(theArray.name, theArray);
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message + "\n" + e.StackTrace);
                    fileLoaded = false;
                }
            }

            return fileLoaded;
        }

        /**
         * Write out the column headers for the MPCArray to the provided worksheet
         **/
        private void outputColHeaders(MPCArray theArray, Excel.Worksheet ws)
        {
            int colNum = -1;
            try
            {
                colNum = 0;

                foreach (MPCArrayColumn arrCol in theArray.columns)
                {
                    if (arrCol.includeInOutput)
                    {
                        colNum++;
                        if (theArray.outputStyle == "cols")
                        {
                            ws.get_Range(String.Concat("A", colNum), Type.Missing).Value2 = arrCol.outputName;
                            ws.get_Range(String.Concat("A", colNum), Type.Missing).AddComment(arrCol.summary);
                        }
                        else
                        {
                            ws.get_Range(String.Concat(getColName(colNum), "1"), Type.Missing).Value2 = arrCol.outputName;
                            ws.get_Range(String.Concat(getColName(colNum), "1"), Type.Missing).AddComment(arrCol.summary);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        /**
         * Convert a column number to a textual column name
         **/
        private string getColName(int colNum)
        {
            string letterList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            string colName = "";
            while (colNum > 0)
            {
                colName = String.Concat(colName, letterList[(colNum % 26) - 1]) ;
                colNum = (int)(colNum / 26);
            }

            return colName;
        }

        /**
         * Skip to the end of an array in a MedPC data file
         **/
        private void skipArray()
        {
            Regex arrayLineRegex = new Regex(@"^\s*(\d+):\s*(-?[0-9\.]+)\s*(?:(-?[0-9\.]+)\s*)?(?:(-?[0-9\.]+)\s*)?(?:(-?[0-9\.]+)\s*)?(?:(-?[0-9\.]+)\s*)?$", RegexOptions.IgnoreCase);

            while (stream.EndOfStream == false && arrayLineRegex.IsMatch(line))
            {
                line = stream.ReadLine();
            }
        }
    }
}