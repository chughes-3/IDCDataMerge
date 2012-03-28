using System;
using System.Windows.Forms;
using System.Text;

namespace InventoryDataMerge2012
{
    public partial class UserWBookWSheet : Form
    {
        InventoryWorkBookClass invSpdSheet;
        public UserWBookWSheet()
        {
            InitializeComponent();
            invSpdSheet = new InventoryWorkBookClass(ref listWBooks, ref listWSheets);  //Links to Excel. Sets up listboxes for viewing
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            invSpdSheet.DisposeX();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            invSpdSheet.SetupRange();   //sets up start and end row values
            invSpdSheet.MergeIDCRows(); //Does the actual work. 
            DialogResult dlg;
            do
            {
                StringBuilder str = new StringBuilder("The IDC Data merge is complete with the following results\r\rMerged IDC records : " + invSpdSheet.rowsMerged.ToString());
                if (invSpdSheet.rowsIdentical != 0)
                    str.Append("\rIDC records discarded due to being duplicates of existing data : " + invSpdSheet.rowsIdentical.ToString());
                if (invSpdSheet.rowsXMLrecsImported != 0)
                    str.Append("\rIDC data records imported from XML file : " + invSpdSheet.rowsXMLrecsImported.ToString());
                if (invSpdSheet.rowsXMLRecsIdentical != 0)
                    str.Append("\rIDC data records NOT imported from XML file -\r    (duplicates of existing spreadsheet rows) : " + invSpdSheet.rowsXMLRecsIdentical.ToString());
                str.Append("\r\r   Merge one or more IDC data files (TaxAideInv2012.xml)?");

                dlg = MessageBox.Show(invSpdSheet.winWrap4MsgBox, str.ToString(), "IDC Data Merge", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (dlg == DialogResult.Cancel || dlg == DialogResult.No)
                {
                    invSpdSheet.Dispose();
                    this.Close();
                    return;
                }
                invSpdSheet.rowsMerged = 0;
                invSpdSheet.rowsXMLRecsIdentical = 0;
                invSpdSheet.rowsXMLrecsImported = 0;
                InvXMLFile xmlFile = new InvXMLFile(invSpdSheet);
                xmlFile.GetIDCXmlData();
                //invSpdSheet.GetMRSerial();  //read in mrserial nos from spreadsheet
                xmlFile.xmlData2End();
                invSpdSheet.MergeIDCRows(); //Does the actual work. 
            } while (dlg == DialogResult.Yes);
            invSpdSheet.Dispose();  //nulls Excel COM interface
            this.Close();
        }

        private void listWSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (invSpdSheet != null)
            {
                invSpdSheet.UpdateWSheet(listWSheets.SelectedIndex);
            }
        }

        private void listWBooks_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (invSpdSheet != null)    //will be null when form is first created
            {
                invSpdSheet.UpdateWBook(listWBooks.SelectedIndex);  //Change workbook selected in invSpdSheet Object
                invSpdSheet.UpdateLbSheets(ref listWSheets);    //WorkBook changed therefore update worksheets
                this.Update();
            }
        }
    }
}
