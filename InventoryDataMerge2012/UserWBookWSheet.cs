using System;
using System.Windows.Forms;

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
            invSpdSheet.Dispose();
            Environment.Exit(1);
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            invSpdSheet.SetupRange();   //sets up start and end row values
            invSpdSheet.MergeIDCRows(); //Does the actual work. 
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
