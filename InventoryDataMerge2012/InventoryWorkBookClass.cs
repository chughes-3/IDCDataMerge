using System;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Linq;

namespace InventoryDataMerge2012
{
    class InventoryWorkBookClass
    {
        const int colMfgSer = 7;    //MUST BE CHANGED IF MFG_SERIAL_NUM Column is changed
        const string colMfgSerG = "G";
        const int colAssTag = 2;    //MUST BE CHANGED IF Asset_Tag column changes!!
        const string colAssTagB = "B";
        const string colMRedSerS = "S";
        const string colMRedSerHdr = "MR_Serial_Number"; //spec'd here to make any col name change obvious
        const string colMfgSerHdr = "Mfg_Serial_Number";
        const string colAssTagHdr = "Asset_Tag";
        const int colIDCEquality = 10;  //used in proc that checks existing IDC data against new idc data. Will need to change if change spreadsheet
        char[] alpha = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'X', 'Y', 'Z' }; //4 convert A1 to R1C1
        internal WindowWrapper winWrap4MsgBox;
        Excel.Application xlApp;
        Excel.Workbooks xlWBooks;
        Excel.Workbook xlWBook = null;
        Excel.Sheets xlWSheets;
        Excel.Worksheet xlWsheet = null;
        List<RowData> rowList = new List<RowData>() { new RowData() { lAssTag = "", lMfgSerNum = "", lMRedSerNum = "" } };  //initial entry to make indexing = excel indexing
        class RowData
        {
            public string lAssTag;
            public string lMfgSerNum;
            public string lMRedSerNum;
        }
        int rowStart = 0;
        int rowEnd;
        int row2Merge;
        int row2bMerged;
        internal int rowsIdentical;
        internal int rowsMerged;
        internal int rowsXMLRecsIdentical;
        internal int rowsXMLrecsImported;

        #region Initialisation, obtain Spreadsheet names,sheets etc, initialise List boxes, update Listboxes
        internal InventoryWorkBookClass()
        {
            try
            {
                xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Excel.Application;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                MessageBox.Show("The Excel application is not open.\rPlease start Excel with the Tax-Aide Inventory Workbook open.\r\r" + e.Message, "IDC Merge");
                Environment.Exit(1);
            }
            Process[] procs = Process.GetProcessesByName("Excel");
            if (procs.Length != 0)
            {
                IntPtr hwnd = procs[0].MainWindowHandle;
                winWrap4MsgBox = new WindowWrapper(hwnd);
            }
        }
        internal InventoryWorkBookClass(ref ListBox lbWBooks, ref ListBox lbWSheets)
            : this()
        {
            xlWBooks = xlApp.Workbooks;
            if (xlWBooks.Count != 0)
            {
                foreach (Excel.Workbook wBook in xlWBooks)
                {
                    lbWBooks.Items.Add(wBook.Name);
                }
                var qry = lbWBooks.Items.Cast<string>().FirstOrDefault(it => Regex.IsMatch(it, ".*inventory.*", RegexOptions.IgnoreCase));
                if (qry != null)
                {
                    lbWBooks.SetSelected(lbWBooks.Items.IndexOf(qry), true);
                    xlWBook = xlWBooks[lbWBooks.Items.IndexOf(qry) + 1];
                }
                else
                {
                    xlWBook = xlWBooks[1];
                    lbWBooks.SetSelected(0, true);
                }
                UpdateLbSheets(ref lbWSheets);
            }
            else
            {
                MessageBox.Show("There are no Excel Workbooks open.\rPlease open the Tax-Aide Inventory Workbook", "IDC Merge");
                xlApp = null;
                Environment.Exit(1);
            }
        }

        internal void UpdateWBook(int indx)
        {
            xlWBook = xlWBooks[indx + 1];
        }

        internal void UpdateWSheet(int indx)
        {
            xlWsheet = xlWSheets[indx + 1];
        }
        internal void UpdateLbSheets(ref ListBox lbWSheets)
        {
            lbWSheets.Items.Clear();
            xlWSheets = xlWBook.Sheets;
            foreach (Excel.Worksheet wsht in xlWSheets)
            {
                lbWSheets.Items.Add(wsht.Name);
            }
            int qry = 0;
            try
            {
                qry = lbWSheets.Items.Cast<string>().Select((item, i) => new { Ite = item, index = i }).FirstOrDefault(it => Regex.IsMatch(it.Ite, ".*inventory.*", RegexOptions.IgnoreCase)).index;
            }
            catch (Exception) { }
            lbWSheets.SetSelected(qry, true);
            xlWsheet = xlWSheets[qry + 1];
        }
        #endregion

        #region  Process Spreadsheet, obtain start,end, Find Matches, merge rows
        internal void SetupRange()
        {
            xlWBook.Activate();
            xlWsheet.Activate();
            Excel.Range stateSearchRng = xlWsheet.Range["A1:A40"];  //40 rows should be enough to find the State
            object[,] stateSearchObj = new object[40, 1];
            stateSearchObj = stateSearchRng.Value2;
            for (int i = 1; i < 40; i++)    //40 rows should be enough to find the State
            {
                if (stateSearchObj[i, 1] != null && stateSearchObj[i, 1].ToString() == "State")
                //string cellValue = (xlWsheet.Cells[i, 1].Value != null) ? xlWsheet.Cells[i, 1].Value.ToString() : "";
                //if (cellValue == "State")
                {
                    rowStart = i + 1;
                    break;
                }
            }
            if (rowStart == 0)
            {
                MessageBox.Show("Unable to find \"State\" column label in the first column\r\rExiting", "IDC Merge");
                xlApp = null;
                Environment.Exit(1);
            }
            string colHeadAss = "";
            string colHeadSer = "";
            try
            {//This is first place access a cell and likely will give issues if spreadsheet in edit mode.
                xlWsheet.Range["A1"].Copy();    //Here to resolve issue if user has left copy or cut selected. Prgram takes control
                xlWsheet.Range["A1"].PasteSpecial();
                colHeadAss = (xlWsheet.Cells[rowStart - 1, colAssTag].Value != null) ? xlWsheet.Cells[rowStart - 1, colAssTag].Value.ToString() : "";
                colHeadSer = (xlWsheet.Cells[rowStart - 1, colMfgSer].Value != null) ? xlWsheet.Cells[rowStart - 1, colMfgSer].Value.ToString() : "";
            }
            catch (Exception)
            {
                MessageBox.Show(winWrap4MsgBox, "The spreadsheet is not accepting programmatic input.\rThe simplest way to fix this is to start the program using a freshly opened spreadsheet in which no editing has been done.\r\r   Exiting!", "IDC Data Merge");
                DisposeX();
            }
            if (colHeadAss != "Asset_Tag" || colHeadSer != "Mfg_Serial_Number")
            {
                MessageBox.Show("The Asset Tag and/or Mfg Serial Number column headings are not in the expected places.\rIs the program pointed at a correctly formatted spreadsheet?", "IDC Merge");
                xlApp = null;
                Environment.Exit(1);
            }
            rowEnd = xlWsheet.Cells.Find("*", SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlPrevious).Row;
            Excel.Range rowsAssTag = xlWsheet.Range[colAssTagB + "1:" + colAssTagB + rowEnd.ToString()];  //start at 1 to keep indexing same as spreadsheet
            object[,] rowsAssTagObj = new object[rowsAssTag.Count, 1];
            Excel.Range rowsMfgSer = xlWsheet.Range[colMfgSerG + "1:" + colMfgSerG + rowEnd.ToString()];  //start 
            object[,] rowsMfgSerObj = new object[rowsMfgSer.Count, 1];
            Excel.Range rowsMred = xlWsheet.Range[colMRedSerS + "1:" + colMRedSerS + rowEnd.ToString()];  //start at 1 to keep indexing same as spreadsheet
            object[,] rowsMredObj = new object[rowsMred.Count, 1];
            rowsAssTagObj = rowsAssTag.Value2;
            rowsMfgSerObj = rowsMfgSer.Value2;
            rowsMredObj = rowsMred.Value2;
            for (int i = 1; i < rowsAssTag.Count + 1; i++)
            {
                rowList.Add(new RowData() { lAssTag = (rowsAssTagObj[i, 1] != null) ? rowsAssTagObj[i, 1].ToString().Trim() : "" });
                rowList[i].lMfgSerNum = ((rowsMfgSerObj[i, 1] != null) ? rowsMfgSerObj[i, 1].ToString().Trim() : "");
                rowList[i].lMRedSerNum = ((rowsMredObj[i, 1]!= null) ? rowsMredObj[i, 1].ToString().Trim() : "");
            }
            xlWsheet.Rows[rowStart].Select();
        }

        internal void MergeIDCRows()
        {
            int currRow = rowStart;
            int matchRow = rowStart + 1;
            int matchedRow;
            int ret;
            while (currRow < rowEnd)
            {
                xlWsheet.Rows[currRow + 3].Select();
                xlWsheet.Rows[currRow].Select();    //force visibility for user
                if (xlApp.WorksheetFunction.CountA(xlWsheet.Rows[currRow]) == 0)
                {
                    xlWsheet.Rows[currRow].Delete();
                    rowList.RemoveAt(currRow);
                    rowEnd--;
                }
                matchedRow = 0;
                ret = FindMatches(out matchedRow, currRow);
                if (ret == 0)
                {// ret ==0 means we have a match, other return of 1 = no match, we have exited on error
                    // Test for error conditions around Machine read data
                    if (CheckMachRedData(currRow, matchedRow) == 0)
                    {//0 means continue regular merge, 1 means identical data row deleted if  otherwise exit on error
                        xlWsheet.Rows[currRow + 1].Insert();
                        xlWsheet.Rows[matchedRow + 1].Cut(xlWsheet.Rows[currRow + 1]);
                        xlWsheet.Rows[matchedRow + 1].Delete();
                        rowList.Insert(currRow + 1, rowList[matchedRow]);
                        rowList.RemoveAt(matchedRow + 1);
                        //Test for error conditions in first cell of rows
                        if (xlWsheet.Cells[currRow, 1].Value == null && xlWsheet.Cells[currRow + 1, 1].Value == null)
                        {
                            MessageBox.Show(winWrap4MsgBox, string.Format("There is an error in row {0} or in row {1}\r\rThe error might be a duplicate Asset_Tag or Mfg_Serial_Num entry\rThe error might be an incorrect previous year data entry\r\r        The row with data from the previous year MUST have data in the first cell\r            - Typically the State or District\r\n        The IDC row first cell must be empty.", currRow, currRow + 1), "IDC Data Merge");
                            xlApp = null;
                            Environment.Exit(1);
                        }
                        if (xlWsheet.Cells[currRow, 1].Value != null && xlWsheet.Cells[currRow + 1, 1].Value != null)
                        {
                            MessageBox.Show(winWrap4MsgBox, string.Format("There is an error in row {0} or in row {1}\r\rThe error might be a duplicate Asset_Tag or Mfg_Serial_Num entry\rThe error might be an incorrect IDC data entry\r\n\r\n        The IDC row first cell must be empty.\r\n        The row with data from the previous year MUST have data in the first cell\r\n           - Typically the State or District", currRow, currRow + 1), "IDC Data Merge");
                            xlApp = null;
                            Environment.Exit(1);
                        }
                        //If user has sorted before running macro need to figure out which row to merge
                        if (xlWsheet.Cells[currRow, 1].Value == null)
                        {//this is 2 merge ie State column is blank
                            row2Merge = currRow;
                            row2bMerged = currRow + 1;
                        }
                        else
                        {
                            row2Merge = currRow + 1;
                            row2bMerged = currRow;
                        }
                        DialogResult dlgResult = MessageBox.Show(winWrap4MsgBox, string.Format(" A Match!\r Row {0} will be merged into Row {1}", row2Merge, row2bMerged), "IDC Merge", MessageBoxButtons.OKCancel);
                        if (dlgResult == DialogResult.Cancel)
                        {
                            xlApp = null;
                            Environment.Exit(0);
                        }
                        xlWsheet.Rows[row2Merge].Copy();
                        xlWsheet.Rows[row2bMerged].PasteSpecial(SkipBlanks: true);
                        xlWsheet.Rows[row2Merge].Delete();
                        if (rowList[row2Merge].lAssTag != "")
                            rowList[row2bMerged].lAssTag = rowList[row2Merge].lAssTag;
                        if (rowList[row2Merge].lMfgSerNum != "")
                            rowList[row2bMerged].lMfgSerNum = rowList[row2Merge].lMfgSerNum;
                        if (rowList[row2Merge].lMRedSerNum != "")
                            rowList[row2bMerged].lMRedSerNum = rowList[row2Merge].lMRedSerNum;
                        rowList.RemoveAt(row2Merge);
                        rowEnd--;
                        rowsMerged++;
                    }
                }
                currRow++;
            }
            return;
        }

        private int CheckMachRedData(int rowExistData, int rowIDCData)
        {//0 = continue regular merge, 1 = identical data row deleted, otherwise program will error exit
            int rowMatch2MredData;
            int matchesMRed = SearchCol4Match(rowStart - 1, rowIDCData - 1, out rowMatch2MredData, rowList[rowIDCData].lMRedSerNum, x => x.lMRedSerNum);
            switch (matchesMRed)
            {
                case 0: //no matches so clean to go
                    return 0;
                case 1: // a match so check on same line as asset/serial if not error
                    if (rowMatch2MredData == rowExistData)
                    {//check for data identicality
                        Excel.Range iDCDataRng = xlWsheet.Range["A" + rowIDCData.ToString() + ":AZ" + rowIDCData.ToString()];
                        Excel.Range rowExistDataRng = xlWsheet.Range["A" + rowExistData.ToString() + ":AZ" + rowExistData.ToString()];
                        object[,] iDCDataRngObj = new object[1, iDCDataRng.Count];
                        iDCDataRngObj = iDCDataRng.Value2;
                        object[,] rowExistDataRngObj = new object[1, rowExistDataRng.Count];
                        rowExistDataRngObj = rowExistDataRng.Value2;
                        //List<string> iDCDataRow = new List<string>() {""};
                        //List<string> existRowData = new List<string>() {""};
                        //for (int j = 1; j < iDCDataRng.Count +1; j++)
                        //{
                        //    existRowData.Add((rowExistDataRng.Cells[1, j].Value != null) ? rowExistDataRng.Cells[1, j].Value.ToString().Trim() : "");
                        //    iDCDataRow.Add((iDCDataRng.Cells[1, j].Value != null) ? iDCDataRng.Cells[1, j].Value.ToString().Trim() : "");
                        //}
                        int i = 0;
                        //string valueExist;
                        for (i = colIDCEquality; i < iDCDataRng.Count; i++)
                        {
                            if (iDCDataRngObj[1, i] != null && rowExistDataRngObj[1, i] != null && iDCDataRngObj[1, i].ToString() != rowExistDataRngObj[1, i].ToString())
                                //if (iDCDataRow[i] != "" && iDCDataRow[i] != existRowData[i])
                                //valueExist = (rowExistDataRng.Cells[1, i].Value != null) ? rowExistDataRng.Cells[1, i].Value.ToString().Trim() : "";
                                //if (iDCDataRng.Cells[1, i].Value != null && iDCDataRng.Cells[1, i].Value.ToString().Trim() != valueExist)
                                break;
                        }
                        if (i != iDCDataRng.Count)
                        {
                            DialogResult dlg1 = MessageBox.Show(winWrap4MsgBox, "The program is merging IDC data row: " + rowIDCData.ToString() + "\rThere is a match based on Asset Tag or Machine Serial No. on row: " + rowExistData.ToString() + "\rThis row already has IDC data, however some of the row: " + rowIDCData.ToString() + " IDC data is different than the IDC data on row: " + rowExistData.ToString() + "\r\rOverwrite the row " + rowExistData.ToString() + " IDC data with the row " + rowIDCData.ToString() + " IDC data?\r\r\t\tYes = Overwrite\r\t\tNo = Keep the existing spreadsheet data\r\t\tCancel = Exit from the program", "IDC Merge", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Error);
                            switch (dlg1)
                            {
                                case DialogResult.Yes:
                                    return 0;
                                case DialogResult.No:
                                    break;
                                case DialogResult.Cancel:
                                    DisposeX();
                                    throw new Exception("Unreachable Code Theoretically");
                            }
                        }
                        //dialog result = NO OR identical data.
                        xlWsheet.Rows[rowIDCData].Delete();
                        rowList.RemoveAt(rowIDCData);
                        rowsIdentical++;
                        rowEnd--;
                        return 1;
                    }
                    else
                    {//error message
                        MessageBox.Show(winWrap4MsgBox, "The program is merging IDC data row: " + rowIDCData.ToString() + "\rThere is a match based on Asset Tag or Machine Serial No. on row: " + rowExistData.ToString() + "\rHowever this Machine Read Serial No. (MR_Serial_Number) has already been used on row: " + rowMatch2MredData.ToString() + "\rEach MR_Serial_Number entry MUST be unique in the spreadsheet! \rAn error in Asset Tag or Manufacturer Serial No. Spreadsheet data entry perhaps??\r\rThe program cannot continue", "IDC Data Merge");
                        DisposeX();
                        throw new Exception("Supposedly Unreachable Code");
                    }
                    throw new Exception("Supposedly Unreachable Code Mach Read testing");
                default:
                    MessageBox.Show(winWrap4MsgBox, "The program is merging IDC data row: " + rowIDCData.ToString() + "\rThere is a match based on Asset Tag or Machine Serial No. on row: " + rowExistData.ToString() + "\rHowever multiple matches for the Machine Read Serial No. (MR_Serial_Number) exist; starting on row: " + rowMatch2MredData.ToString() + "\rEach MR_Serial_Number entry MUST be unique in the spreadsheet! \r\rThe program cannot continue", "IDC Data Merge");
                    DisposeX();
                    throw new Exception("Supposedly Unreachable Code");
            }
        }

        private int FindMatches(out int matchedRow, int rowCurrent)
        {
            int matchedRowNoSerial;
            int matchedRowNoAsset;
            int matchCountSerial = 0;
            int matchCountAsset = 0;
            matchedRow = 0;
            DialogResult dlgResult;

            string currRowAssTag = rowList[rowCurrent].lAssTag;
            string currRowSerTag = rowList[rowCurrent].lMfgSerNum;
            matchCountSerial = SearchCol4Match(rowCurrent, rowEnd, out matchedRowNoSerial, currRowSerTag, x => x.lMfgSerNum);
            matchCountAsset = SearchCol4Match(rowCurrent, rowEnd, out matchedRowNoAsset, currRowAssTag, x => x.lAssTag);
            switch (matchCountSerial)
            {
                case 0:
                    switch (matchCountAsset)
                    {
                        case 0:
                            return 1;   //no match
                        case 1:         //good match
                            matchedRow = matchedRowNoAsset;
                            return 0;
                        default:        //multiple matches
                            MessageBox.Show(winWrap4MsgBox, String.Format("The program is finding matches for row {2}\r\rThere are {0} matches for Asset_Tag = {1}\rOne of which is on row: {3}\r\rThe programmatic merge cannot continue", matchCountAsset, currRowAssTag, rowCurrent, matchedRowNoAsset), "IDC Merge");
                            DisposeX();
                            throw new Exception("Supposedly Unreachable Code");
                    }
                case 1:
                    switch (matchCountAsset)
                    {
                        case 0:
                            matchedRow = matchedRowNoSerial;
                            return 0;
                        case 1:
                            if (matchedRowNoSerial == matchedRowNoAsset)
                            {
                                matchedRow = matchedRowNoAsset;
                                return 0;
                            }
                            else
                            {
                                MessageBox.Show(winWrap4MsgBox, String.Format("The program is finding matches for row {2}\r\rThere is a match on row {0} for Mfg_Serial_Num = {1}\rThere is a match on row {3} for Asset_Tag = {4}\r\rThese are conflicting!!\rThe programmatic merge cannot continue", matchedRowNoSerial, currRowSerTag, rowCurrent, matchedRowNoAsset, currRowAssTag), "IDC Merge");
                                DisposeX();
                                throw new Exception("Supposedly Unreachable Code");
                            }
                        default:
                            dlgResult = MessageBox.Show(winWrap4MsgBox, String.Format("The program is finding matches for row {2}\r\rThere are {0} matches for Asset_Tag = {1}\rThere is a match on row {3} for Mfg_Serial_Num = {4}\r\rContinue the merge with the Mfg_Serial_Num match?", matchCountAsset, currRowAssTag, rowCurrent, matchedRowNoSerial, currRowSerTag), "IDC Merge", MessageBoxButtons.OKCancel);
                            if (dlgResult == DialogResult.OK)
                            {
                                matchedRow = matchedRowNoSerial;
                                return 0;
                            }
                            else
                            {
                                DisposeX();
                                throw new Exception("Supposedly Unreachable Code");
                            }
                    }
                default:
                    if (matchCountAsset == 1)
                    {
                        dlgResult = MessageBox.Show(winWrap4MsgBox, String.Format("The program is finding matches for row {2}\r\rThere are {0} matches for Mfg_Serial_Num = {1}\rThere is a match on row {3} for Asset_Tag = {4}\r\rContinue the merge with the Asset_Tag match?", matchCountSerial, currRowSerTag, rowCurrent, matchedRowNoAsset, currRowAssTag), "IDC Merge", MessageBoxButtons.OKCancel);
                        if (dlgResult == DialogResult.OK)
                        {
                            matchedRow = matchedRowNoAsset;
                            return 0;
                        }
                        else
                        {
                            DisposeX();
                            throw new Exception("Supposedly Unreachable Code");
                        }
                    }
                    else
                    {
                        MessageBox.Show(winWrap4MsgBox, String.Format("The program is finding matches for row {2}\r\rThere are {0} matches for Asset_Tag = {1}\rThere are {3} matches for Mfg_Ser_Num = {4}\rThe programmatic merge cannot continue", matchCountAsset, currRowAssTag, rowCurrent, matchCountSerial, currRowSerTag), "IDC Merge");
                        DisposeX();
                        throw new Exception("Supposedly Unreachable Code");
                    }
            }
        }

        #endregion

        private int SearchCol4Match(int startRow, int endRow, out int matchedRow, string currRowVal, Func<RowData, string> gdata)
        {//returns number of matches and the last row matched (in out)
            int matchCount = 0;
            matchedRow = 0;
            string pattern = ".*\\d.*"; //looks for a single digit anywhere in the string
            if (Regex.IsMatch(currRowVal, pattern)) //ass tag may have n/a when serial is ok or vice versa in which case just return
            {
                for (int k = startRow + 1; k < endRow + 1; k++)
                {
                    if (currRowVal == gdata(rowList[k]))
                    {
                        matchedRow = k;
                        matchCount++;
                    }
                }
            }
            return matchCount;
        }

        internal void XLRowFromXml(System.Xml.Linq.XElement iDCSys)
        {
            //check if data already exists
            for (int i = rowStart; i < rowEnd + 1; i++)
            {
                if (rowList[i].lMRedSerNum == iDCSys.Element(colMRedSerHdr).Value.Trim())
                {

                    Excel.Range rowDataRng = xlWsheet.Range["A" + i.ToString() + ":AZ" + i.ToString()];
                    object[,] rowDataObj = new object[1, rowDataRng.Count];
                    rowDataObj = rowDataRng.Value2;
                    int j = 0;
                    for (j = 0; j < iDCSys.Elements().Count(); j++)
                    {
                        if (rowDataObj[1, j + 1] != null && rowDataObj[1, j + 1].ToString() != iDCSys.Elements().ElementAt(j).Value.Trim())
                            break;
                    }
                    if (j == iDCSys.Elements().Count())
                    {
                        rowsXMLRecsIdentical++;
                        return;
                    }
                }
            }
            string colEnd = "";
            rowEnd++;   //we are extending worksheet by one row
            if (iDCSys.Elements().Count() < 26)
                colEnd = alpha[iDCSys.Elements().Count() - 2].ToString();
            else if (iDCSys.Elements().Count() < 52)
                colEnd = "A" + alpha[iDCSys.Elements().Count() - 2].ToString();     //if more than 52 cols will throw an error
            object[,] objData = new object[1, iDCSys.Elements().Count()];
            Excel.Range rngIDC = xlWsheet.Range["A" + rowEnd.ToString() + ":" + colEnd + rowEnd.ToString()];
            for (int i = 0; i < iDCSys.Elements().Count(); i++)
            {
                objData[0, i] = iDCSys.Elements().ElementAt(i).Value.Trim();
            }
            rngIDC.Value2 = objData;
            //Next update rowList
            rowList.Add(new RowData { lAssTag = iDCSys.Element(colAssTagHdr).Value.Trim(), lMfgSerNum = iDCSys.Element(colMfgSerHdr).Value.Trim(), lMRedSerNum = iDCSys.Element(colMRedSerHdr).Value.Trim() });
            rowsXMLrecsImported++;
        }
        internal void DisposeX()
        {
            xlApp = null;
            Environment.Exit(1);
        }
        internal void Dispose()
        {
            xlApp = null;
        }
    }

    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        public IntPtr Handle
        {
            get { return _hwnd; }
        }

        private IntPtr _hwnd;
    }
}
