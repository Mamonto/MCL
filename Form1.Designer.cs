namespace DataEasy
{
    partial class Form1
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
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Items.AddRange(new object[] {
            "1)\tThe EasyMasterClientList application allows users to load any ",
            "\tMicrosoft Access file (file format \'.mdb\') and to select, add, delete, ",
            "\tinsert, and modify the database and then save all the changes to the ",
            "\toriginal Access Database file. ",
            "a)\tTo load the database, select File->Open",
            "b)\tNavigate to the folder where the “MasterClientList2003.mdb “ file was ",
            "\tcopied and select it. Click “Open” button.",
            "c)\tThe updates can be made in the Tex Boxes or the spreadsheet bellow.",
            "d)\tOnce the changes are made, Click “Update” button and select File->Save.",
            "e)\tTo close the application, select File->Exit  ",
            "",
            "2)\tThe Search Functionality can be invoked from the toolbar above ",
            "\tSearch->Search Data. By right clicking the search form’s column headers, ",
            "\tthe user can filter in the desired element. Multiple filtering can be done by",
            "\tfiltering more than one column. A simple search can be applied by typing in the " +
                "search textbox to find data in the database.   ",
            "Note: Updates are not available in the Search window."});
            this.listBox1.Location = new System.Drawing.Point(0, 0);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(367, 290);
            this.listBox1.TabIndex = 0;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 293);
            this.Controls.Add(this.listBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox listBox1;

    }
}