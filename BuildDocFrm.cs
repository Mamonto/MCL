using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

using System.Data.OleDb;
using System.Collections;
using Utility.NiceMenu;
using Microsoft.Office;


namespace DataEasy
{
    public class BuildDocFrm : System.Windows.Forms.Form
    {
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGrid dGrid;
        private DataSet accessDataSet = new DataSet();
        private System.Windows.Forms.ComboBox comboTables;
        private Label label1;
        private ListBox listBox1;
        private DataTable accessDataTable;

        public BuildDocFrm(Form motherFrm, DataSet dataset, DataGrid datagrid, ComboBox combo_tables)
        {
            InitializeComponent();

            accessDataSet = dataset;
            dGrid = datagrid;
            comboTables = combo_tables;

            //accessDataTable = accessDataSet.Tables[tableName];
            accessDataTable = accessDataTable = accessDataSet.Tables[comboTables.Text];

            //Find all columns and put them in the listBox
            string temp1_get_header_from_db_table = "";
            for (int i = 0; i < accessDataTable.Columns.Count; i++)
            {
                //Create a bookmark name from the column name of the database tablle,
                //modify it to the bookmark alike name.
                //and display in the listBox1 of this Form
                temp1_get_header_from_db_table = accessDataTable.Columns[i].Caption;
                temp1_get_header_from_db_table = "BM_" + temp1_get_header_from_db_table.Replace(" ", "_");
                listBox1.Items.Add(temp1_get_header_from_db_table);
            }
        }
       

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BuildDocFrm));
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(266, 23);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(136, 43);
            this.button1.TabIndex = 0;
            this.button1.Text = "Build with Bookmarks";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.buttonBuildDoc_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(23, 23);
            this.label1.MaximumSize = new System.Drawing.Size(300, 250);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(228, 189);
            this.label1.TabIndex = 1;
            this.label1.Text = resources.GetString("label1.Text");
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 16;
            this.listBox1.Location = new System.Drawing.Point(26, 228);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(223, 148);
            this.listBox1.TabIndex = 4;
            // 
            // BuildDocFrm
            // 
            this.ClientSize = new System.Drawing.Size(414, 399);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "BuildDocFrm";
            this.Text = "Document Builder";
            this.ResumeLayout(false);

        }

        private void buttonBuildDoc_Click(object sender, System.EventArgs e)
        {
            try
            {
                //Get all the System.DataTypes of all the columns in the table and assign them to the array colType
                string[] colType = new string[accessDataTable.Columns.Count];

                //Get the values from the database and copy to new variavles
                int current_row_number = (int)dGrid.CurrentRowIndex;

                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                //Start Word and create a new document.
                Word._Application oWord;
                Word._Document oDoc;
                oWord = new Word.Application();
                oWord.Visible = true;

                
                //OBJECTS OF FALSE AND TRUE
                Object oTrue = true;
                Object oFalse = false;
  
                
                //THE LOCATION OF THE TEMPLATE FILE ON THE MACHINE
                //Object oTemplatePath = "C:\\Program Files\\MyTemplate.dot";
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.FileName = "";
                //Make sure only MS Word Doc files can be opened
                //by using a filter
                openFile.Filter = "Microsoft Word Application (*.doc)|*.doc";    

                System.Windows.Forms.DialogResult res = openFile.ShowDialog();
                if (res == System.Windows.Forms.DialogResult.Cancel) return;

                Object oTemplatePath = openFile.FileName;

                //ADDING A NEW DOCUMENT FROM A TEMPLATE
                oDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

                //Find Total number of pages in the template document
                //Object TotalPages = Word.WdFieldType.wdFieldNumPages;
                //oWord.ActiveWindow.Selection.Fields.Add(oWord.Selection.Range, ref TotalPages, ref oMissing, ref oMissing);
                //MessageBox.Show(TotalPages.ToString() );

                //Iterate through bookmarks in a document template
                for (int j = 1; j <= oDoc.Bookmarks.Count; j++)
                {
                    object objI = j;
                    //Here is your bookmark found in a document template.
                    //Store it where ever you want:
                    //MessageBox.Show(oDoc.Bookmarks.get_Item(ref objI).Name);
                    //object found_bookname_in_template = oDoc.Bookmarks.get_Item(ref objI).Name;
                    string found_bookname_in_template = oDoc.Bookmarks.get_Item(ref objI).Name.ToString();

                    //If the bookmark has extention "_2", remove it.
                    //Note: Sometimes 2 or more bookmarks refer to the same data in DB to place that 
                    //      same data in different places in a document. This code handles extentions 
                    //      up to "_9".
                    string without_last_twocharacters = found_bookname_in_template.Substring(0, found_bookname_in_template.Length-2);

                    //Create dynamically all the bookmarks from the database headers
                    for (int i = 0; i < accessDataTable.Columns.Count; i++)
                    {

                        colType[i] = accessDataTable.Columns[i].DataType.ToString();
                        //listBox1.Text = accessDataTable.Columns[i].Caption;

                        //Keep a copy of the current column's header/title
                        string temp_column_name = accessDataTable.Columns[i].Caption;

                        //Create a bookmark name from the column name of the database table
                        string temp_bookmark_from_db_table = "BM_" + temp_column_name.Replace(" ", "_");
                        //comboBox1.Text = temp_bookmark_from_db_table;

                        if (temp_bookmark_from_db_table == found_bookname_in_template || temp_bookmark_from_db_table == without_last_twocharacters)
                        {
                            string this_cell_data = this.accessDataSet.Tables[comboTables.Text].Rows[current_row_number][temp_column_name].ToString();
                            object oBookMark = found_bookname_in_template;
                            oDoc.Bookmarks.get_Item(ref oBookMark).Range.Text = this_cell_data;
                        }
                    }
                }
                
            }
            //Catch all the erros and report them
            catch (Exception)
            {
                MessageBox.Show("Unfortunately this document cannot be created. Please check bookmarks.", "WARNING!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.Close();
            }



            

            //Close this form.
            //this.Close();
        }

             
    }
}
