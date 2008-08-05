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
        string msn;
        System.Collections.Hashtable xmlVars;
        System.Collections.Hashtable xmlArrays;

        public FileParser(string theFilename, Excel.Workbook theWorkbook)
        {
            stream = new System.IO.StreamReader(theFilename);
            wb = theWorkbook;
        }

        public FileParser(System.IO.StreamReader theStream, Excel.Workbook theWorkbook)
        {
            stream = theStream;
            wb = theWorkbook;
        }

        public void Parse()
        {
            xmlVars = new System.Collections.Hashtable();
            xmlArrays = new System.Collections.Hashtable();

            parseHeader();
            readXml();
            parseVariables();
        }

        private void parseHeader()
        {
            Excel.Worksheet wsOverview = (Excel.Worksheet) wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Excel.XlSheetType.xlWorksheet);
            wsOverview.Name = "Overview";

            int colonPos = 0, rowNum = 1;
            string label, value;

            string line;

            while ((line = stream.ReadLine()) != null)
            {
                colonPos = line.IndexOf(':');
                if (colonPos != -1)
                {
                    label = line.Substring(0, colonPos);
                    value = line.Substring(colonPos + 1, line.Length - colonPos - 1).Trim();
                    wsOverview.get_Range(String.Concat('A', rowNum.ToString()), Type.Missing).Value2 = label;
                    wsOverview.get_Range(String.Concat('B', rowNum.ToString()), Type.Missing).Value2 = value;
                    rowNum++;
                    if (label.Equals("MSN"))
                    {
                        msn = value;
                        break;
                    }
                }
            }
        }

        private void parseVariables()
        {
            try
            {
                Excel.Worksheet wsVars = null;

                int rowNum = 1;
                string label, value;
                value = "";

                Regex varStartRegex = new Regex(@"^\s*([A-Za-z]):\s*(?:([0-9\.]+)\s*)?$", RegexOptions.IgnoreCase);

                string line;
                line = stream.ReadLine();
                while (line != null)
                {
                    if (varStartRegex.IsMatch(line)) //match -> a variable name (and possibly value) is present on this line
                    {
                        Match theMatch = varStartRegex.Match(line);
                        label = theMatch.Groups[1].Captures[0].Value;

                        if (theMatch.Groups[2].Captures.Count > 0) //if >0 then variable value on same line. Otherwise it is an array.\
                        {
                            label = theMatch.Groups[1].Value;
                            value = theMatch.Groups[2].Value;

                            if (xmlVars.ContainsKey(label))
                            {
                                if (wsVars == null)
                                {
                                    wsVars = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Excel.XlSheetType.xlWorksheet);
                                    wsVars.Name = "Variables";
                                }
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
                            if (xmlArrays.ContainsKey(label))
                            {
                                readDataArray(ref line, label);
                            }
                            else
                            {
                                line = stream.ReadLine();
                            }
                        }
                    }
                    else
                    {
                        line = stream.ReadLine();
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + e.StackTrace);
            }
        }

        private void readDataArray(ref string line, string label)
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

            Regex arrayValuesRegex = new Regex(@"^\s*(\d+):\s*(-?[0-9\.]+)\s*(?:(-?[0-9\.]+)\s*)?(?:(-?[0-9\.]+)\s*)?(?:(-?[0-9\.]+)\s*)?(?:(-?[0-9\.]+)\s*)?$", RegexOptions.IgnoreCase);

            line = stream.ReadLine();
            while (line != null && arrayValuesRegex.IsMatch(line))
            {
                Match theMatch = arrayValuesRegex.Match(line);
                for (int i = 2; i < theMatch.Groups.Count; i++)
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

        private void readXml()
        {
            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
            try
            {
                xmlDoc.Load(String.Concat("c:\\MED-PC IV\\MPC\\", msn, ".MPC.xml"));
            }
            catch (System.IO.FileNotFoundException fnfe)
            {
                OpenFileDialog theDialog = new OpenFileDialog();
                theDialog.InitialDirectory = "c:\\MED-PC IV\\MPC";
                //theDialog.RestoreDirectory = true;

                if (theDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        xmlDoc.Load(theDialog.FileName);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message + "\n" + e.StackTrace);
                        return;
                    }
                }
            }

            try
            {
                System.Xml.XmlNodeList nodes = xmlDoc.SelectNodes("/program/variable");
                foreach (System.Xml.XmlNode node in nodes)
                {
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
            }
        }

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
    }
}