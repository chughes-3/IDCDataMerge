namespace InventoryDataMerge2012
{
    partial class UserWBookWSheet
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
            this.label1 = new System.Windows.Forms.Label();
            this.listWBooks = new System.Windows.Forms.ListBox();
            this.listWSheets = new System.Windows.Forms.ListBox();
            this.labelWBs = new System.Windows.Forms.Label();
            this.labelWSs = new System.Windows.Forms.Label();
            this.labelConfimB = new System.Windows.Forms.Label();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(327, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Confirm the Excel Wookbook and WorkSheet to be used";
            // 
            // listWBooks
            // 
            this.listWBooks.FormattingEnabled = true;
            this.listWBooks.HorizontalScrollbar = true;
            this.listWBooks.Location = new System.Drawing.Point(16, 49);
            this.listWBooks.Name = "listWBooks";
            this.listWBooks.Size = new System.Drawing.Size(322, 56);
            this.listWBooks.TabIndex = 1;
            this.listWBooks.SelectedIndexChanged += new System.EventHandler(this.listWBooks_SelectedIndexChanged);
            // 
            // listWSheets
            // 
            this.listWSheets.FormattingEnabled = true;
            this.listWSheets.HorizontalScrollbar = true;
            this.listWSheets.Location = new System.Drawing.Point(16, 139);
            this.listWSheets.Name = "listWSheets";
            this.listWSheets.Size = new System.Drawing.Size(322, 69);
            this.listWSheets.TabIndex = 1;
            this.listWSheets.SelectedIndexChanged += new System.EventHandler(this.listWSheets_SelectedIndexChanged);
            // 
            // labelWBs
            // 
            this.labelWBs.AutoSize = true;
            this.labelWBs.Location = new System.Drawing.Point(13, 33);
            this.labelWBs.Name = "labelWBs";
            this.labelWBs.Size = new System.Drawing.Size(133, 13);
            this.labelWBs.TabIndex = 0;
            this.labelWBs.Text = "WorkBooks currently open";
            // 
            // labelWSs
            // 
            this.labelWSs.AutoSize = true;
            this.labelWSs.Location = new System.Drawing.Point(13, 123);
            this.labelWSs.Name = "labelWSs";
            this.labelWSs.Size = new System.Drawing.Size(218, 13);
            this.labelWSs.TabIndex = 0;
            this.labelWSs.Text = "WorkSheets available in selected Workbook";
            // 
            // labelConfimB
            // 
            this.labelConfimB.AutoSize = true;
            this.labelConfimB.Location = new System.Drawing.Point(60, 226);
            this.labelConfimB.Name = "labelConfimB";
            this.labelConfimB.Size = new System.Drawing.Size(203, 26);
            this.labelConfimB.TabIndex = 0;
            this.labelConfimB.Text = "The program will run against the selected \r\n      Excel WorkBook and Worksheet.";
            // 
            // buttonOK
            // 
            this.buttonOK.Location = new System.Drawing.Point(231, 269);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 23);
            this.buttonOK.TabIndex = 2;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(89, 269);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 2;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // UserWBookWSheet
            // 
            this.AcceptButton = this.buttonOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size(357, 315);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.listWSheets);
            this.Controls.Add(this.listWBooks);
            this.Controls.Add(this.labelWSs);
            this.Controls.Add(this.labelWBs);
            this.Controls.Add(this.labelConfimB);
            this.Controls.Add(this.label1);
            this.Name = "UserWBookWSheet";
            this.Text = "IDC Inventory 2012 Merge";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox listWBooks;
        private System.Windows.Forms.ListBox listWSheets;
        private System.Windows.Forms.Label labelWBs;
        private System.Windows.Forms.Label labelWSs;
        private System.Windows.Forms.Label labelConfimB;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCancel;
    }
}

