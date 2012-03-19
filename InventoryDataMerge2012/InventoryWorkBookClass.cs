using System;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace InventoryDataMerge2012
{
    class InventoryWorkBookClass
    {
        const int mfgSerCol = 7;    //MUST BE CHANGED IF MFG_SERIAL_NUM Column is changed
        const int assTagCol = 2;    //MUST BE CHANGED IF Asset_Tag column changes!!
        WindowWrapper winWrap4MsgBox;
        Excel.Application xlApp;
        Excel.Workbooks xlWBooks;
        Excel.Workbook xlWBook = null;
        Excel.Sheets xlWSheets;
        Excel.Worksheet xlWsheet = null;
        int startRow = 0;
        int endRow;
        int row2Merge;
        int row2bMerged;
        internal InventoryWorkBookClass()
        {
            try
            {
                xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Excel.Application;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                MessageBox.Show("The Excel application is not open.\rPlease start Excel with the Tax-Aide Inventory Workbook open.\r\r" + e.Message,"IDC Merge");
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

                foreach (var item in lbWBooks.Items)
                {
                    int indx = item.ToString().IndexOf("Inventory", StringComparison.CurrentCultureIgnoreCase);
                    if (indx > 0)
                    {
                        lbWBooks.SetSelected(lbWBooks.Items.IndexOf(item), true);
                        xlWBook = xlWBooks[lbWBooks.Items.IndexOf(item) + 1];
                        break;
                    }
                }
                if (xlWBook == null)
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

            foreach (var lbShItem in lbWSheets.Items)
            {
                int indx = lbShItem.ToString().IndexOf("Inventory", StringComparison.CurrentCultureIgnoreCase);
                if (indx > 0)
                {
                    lbWSheets.SetSelected(lbWSheets.Items.IndexOf(lbShItem), true);
                    xlWsheet = xlWSheets[lbWSheets.Items.IndexOf(lbShItem) + 1];
                    break;
                }
            }
            if (lbWSheets.SelectedIndex == -1)   //may have not been a selection
            {
                lbWSheets.SetSelected(0, true);
                xlWsheet = xlWSheets[1];
            }
        }
        internal void SetupRange()
        {
            xlWBook.Activate();
            xlWsheet.Activate();
            for (int i = 1; i < 40; i++)    //40 rows should be enough to find the State
            {
                string cellValue = (xlWsheet.Cells[i, 1].Value != null) ? xlWsheet.Cells[i, 1].Value.ToString() : ""; 
                if (cellValue == "State")
                {
                    startRow = i + 1;
                    break;
                }
            }
            if (startRow == 0)
            {
                MessageBox.Show("Unable to find \"State\" column label in the first column\r\rExiting", "IDC Merge");
                xlApp = null;
                Environment.Exit(1);
            }
            endRow = xlWsheet.Cells.Find("*", SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlPrevious).Row;
            //Excel.Range objRange = xlWsheet.Rows[startRow];
            //objRange.Select();
            xlWsheet.Rows[startRow].Select();
        }

        internal void MergeIDCRows()
        {
            int currRow = startRow;
            int matchRow = startRow + 1;
            int matchedRow;
            int ret;
            while (currRow < endRow)
            {
                xlWsheet.Rows[currRow + 3].Select();
                xlWsheet.Rows[currRow].Select();    //force visibility for user
                if (xlApp.WorksheetFunction.CountA(xlWsheet.Rows[currRow]) == 0)
                {
                    xlWsheet.Rows[currRow].Delete();
                    endRow--;
                }
                matchedRow = 0;
                ret = FindMatches(out matchedRow, currRow);
                if (ret == 0)
                {//We have a match, other return of 1 = no match, we have exited on error
                    xlWsheet.Rows[currRow + 1].Insert();
                    xlWsheet.Rows[matchedRow + 1].Cut(xlWsheet.Rows[currRow + 1]);
                    xlWsheet.Rows[matchedRow + 1].Delete();
                    //Test for error conditions in first cell of rows
                    if (xlWsheet.Cells[currRow, 1].Value == null && xlWsheet.Cells[currRow + 1, 1].Value == null)
                    {
                        MessageBox.Show(winWrap4MsgBox, string.Format("There is an error in row {0} or in row {1}\r\rThe error might be a duplicate Asset_Tag or Mfg_Serial_Num entry\rThe error might be an incorrect previous year data entry\r\r        The row with data from the previous year MUST have data in the first cell\r            - Typically the State or District\r\n        The IDC row first cell must be empty.", currRow, currRow + 1), "IDC Merge");
                        xlApp = null;
                        Environment.Exit(1);
                    }
                    if (xlWsheet.Cells[currRow, 1].Value != null && xlWsheet.Cells[currRow + 1, 1].Value != null)
                    {
                        MessageBox.Show(winWrap4MsgBox, string.Format("There is an error in row {0} or in row {1}\r\rThe error might be a duplicate Asset_Tag or Mfg_Serial_Num entry\rThe error might be an incorrect IDC data entry\r\n\r\n        The IDC row first cell must be empty.\r\n        The row with data from the previous year MUST have data in the first cell\r\n           - Typically the State or District", currRow, currRow + 1), "IDC Merge");
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
                    endRow--;
                }
                currRow++;
            }
            MessageBox.Show(winWrap4MsgBox, "The merge is complete.", "IDC Merge");
            return;
        }

        private int FindMatches(out int matchedRow, int currRow)
        {
            int matchedRowNoSerial;
            int matchedRowNoAsset;
            int matchCountSerial = 0;
            int matchCountAsset = 0;
            matchedRow = 0;
            DialogResult dlgResult;

            string currRowAssTag = (xlWsheet.Cells[currRow, assTagCol].Value != null) ? xlWsheet.Cells[currRow, assTagCol].Value.ToString().Trim() : "";
            string currRowSerTag = (xlWsheet.Cells[currRow, mfgSerCol].Value != null) ? xlWsheet.Cells[currRow, mfgSerCol].Value.ToString().Trim() : "";
            matchCountSerial = SearchCol4Match(currRow, out matchedRowNoSerial, mfgSerCol, currRowSerTag);
            matchCountAsset = SearchCol4Match(currRow, out matchedRowNoAsset, assTagCol, currRowAssTag);
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
                            MessageBox.Show(winWrap4MsgBox, String.Format("The program is finding matches for row {2}\r\rThere are {0} matches for Asset_Tag = {1}\r\rThe programmatic merge cannot continue", matchCountAsset, currRowAssTag, currRow), "IDC Merge");
                            xlApp = null;
                            Environment.Exit(1);
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
                                MessageBox.Show(winWrap4MsgBox, String.Format("The program is finding matches for row {2}\r\rThere is a match on row {0} for Mfg_Serial_Num = {1}\rThere is a match on row {3} for Asset_Tag = {4}\r\rThese are conflicting!!\rThe programmatic merge cannot continue", matchedRowNoSerial, currRowSerTag, currRow, matchedRowNoAsset, currRowAssTag), "IDC Merge");
                                xlApp = null;
                                Environment.Exit(1);
                                throw new Exception("Supposedly Unreachable Code");
                            }
                        default:
                            dlgResult = MessageBox.Show(winWrap4MsgBox, String.Format("The program is finding matches for row {2}\r\rThere are {0} matches for Asset_Tag = {1}\rThere is a match on row {3} for Mfg_Serial_Num = {4}\r\rContinue the merge with the Mfg_Serial_Num match?", matchCountAsset, currRowAssTag, currRow, matchedRowNoSerial, currRowSerTag), "IDC Merge", MessageBoxButtons.OKCancel);
                            if (dlgResult == DialogResult.OK)
                            {
                                matchedRow = matchedRowNoSerial;
                                return 0;
                            }
                            else
                            {
                                xlApp = null;
                                Environment.Exit(1);
                                throw new Exception("Supposedly Unreachable Code");
                            }
                    }
                default:
                    if (matchCountAsset == 1)
                    {
                        dlgResult = MessageBox.Show(winWrap4MsgBox, String.Format("The program is finding matches for row {2}\r\rThere are {0} matches for Mfg_Serial_Num = {1}\rThere is a match on row {3} for Asset_Tag = {4}\r\rContinue the merge with the Asset_Tag match?", matchCountSerial, currRowSerTag, currRow, matchedRowNoAsset, currRowAssTag), "IDC Merge", MessageBoxButtons.OKCancel);
                        if (dlgResult == DialogResult.OK)
                        {
                            matchedRow = matchedRowNoAsset;
                            return 0;
                        }
                        else
                        {
                            xlApp = null;
                            Environment.Exit(1);
                            throw new Exception("Supposedly Unreachable Code");
                        }
                    }
                    else
                    {
                        MessageBox.Show(winWrap4MsgBox, String.Format("The program is finding matches for row {2}\r\rThere are {0} matches for Asset_Tag = {1}\rThere are {3} matches for Mfg_Ser_Num = {4}\rThe programmatic merge cannot continue", matchCountAsset, currRowAssTag, currRow, matchCountSerial, currRowSerTag), "IDC Merge");
                        xlApp = null;
                        Environment.Exit(1);
                        throw new Exception("Supposedly Unreachable Code");
                    }
            }
        }

        private int SearchCol4Match(int currRow, out int matchedRow, int searchCol, string currRowVal)
        {//returns number of matches and the last row matched (in out)
            int matchCount = 0;
            string testVal;
            matchedRow = 0;
            string pattern = ".*\\d.*"; //looks for a single digit anywhere in the string
            if (Regex.IsMatch(currRowVal, pattern)) //ass tag may have n/a when serial is ok or vice versa in which case just return
            {
                for (int k = currRow + 1; k < endRow + 1; k++)
                {
                    testVal = (xlWsheet.Cells[k, searchCol].Value != null) ? xlWsheet.Cells[k, searchCol].Value.ToString().Trim() : "";  //In case numbers stored as numbers plus eliminate leading/trailing WS
                    if (currRowVal == testVal)
                    {
                        matchedRow = k;
                        matchCount++;
                    }
                }
            }
            return matchCount;
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
