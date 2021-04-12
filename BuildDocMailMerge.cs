using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Data.OleDb;
using System.Collections;
using Utility.NiceMenu;
using Microsoft.Office;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace DataEasy
{
    public partial class BuildDocMailMerge : System.Windows.Forms.Form
    {

        private System.Windows.Forms.DataGrid dGrid;
        private DataSet accessDataSet = new DataSet();
        private System.Windows.Forms.ComboBox comboTables;
        private DataTable accessDataTable;

        public BuildDocMailMerge(DataSet dataset, DataGrid datagrid, ComboBox combo_tables)
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
                temp1_get_header_from_db_table = "<<" + temp1_get_header_from_db_table.Replace(" ", "_") + ">>";
                listBox1.Items.Add(temp1_get_header_from_db_table);
            }        
        }

        public void MailMerge()
        {
        }

        public void button1_Click(object sender, EventArgs e)
        {
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
            openFile.Filter = "Microsoft Word Application (*.do*)|*.do*";

            System.Windows.Forms.DialogResult res = openFile.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.Cancel) return;

            Object oTemplatePath = openFile.FileName;

            //ADDING A NEW DOCUMENT FROM A TEMPLATE
            oDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);


            ///////////////////////////////////////////////////////
            //OBJECT OF MISSING "NULL VALUE"
            //Object oMissing = System.Reflection.Missing.Value;
            //Object oTemplatePath = "C:\\Law\\Documents\\WordTemplates\\DISCLOSURE LETTER - TORONTO - by FAX.dotx";
            //Object oTemplatePath = "C:\\Users\\mikhailma\\Documents\\MyTemplate.doc";

            //Application wordApp = new Application();

            //Document wordDoc = new Document();

            //oDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            foreach (Word.Field myMergeField in oDoc.Fields)
            {
                Word.Range rngFieldCode = myMergeField.Code;

                String fieldText = rngFieldCode.Text;

                // ONLY GETTING THE MAILMERGE FIELDS
                if (fieldText.StartsWith(" MERGEFIELD"))
                {
                    // THE TEXT COMES IN THE FORMAT OF (with MS Word 2007):
                    // MERGEFIELD  MyFieldName  \\* MERGEFORMAT
                    // THIS HAS TO BE EDITED TO GET ONLY THE FIELDNAME "MyFieldName"

                    //Int32 endMerge = fieldText.IndexOf("\\");
                    //Int32 fieldNameLength = fieldText.Length - endMerge;
                    //String fieldName = fieldText.Substring(11, endMerge - 11);

                    // THE TEXT COMES IN THE FORMAT OF (with MS Word older than 2007):
                    String fieldName = fieldText.Replace("\"", "");
                    fieldName = fieldName.Replace("MERGEFIELD", "");

                    // GIVES THE FIELDNAMES AS THE USER HAD ENTERED IN .dot FILE
                    fieldName = fieldName.Trim();


                    //Find Total number of pages in the template document
                    //Object TotalPages = Word.WdFieldType.wdFieldNumPages;
                    //oWord.ActiveWindow.Selection.Fields.Add(oWord.Selection.Range, ref TotalPages, ref oMissing, ref oMissing);
                    //MessageBox.Show(TotalPages.ToString() );

                    
                    //Get all the headers from the database and create 
                    //the Mail Merge Fields From them
                    for (int i = 0; i < accessDataTable.Columns.Count; i++)
                    {
                        //colType[i] = accessDataTable.Columns[i].DataType.ToString();

                        //Keep a copy of the current column's header/title
                        string temp_column_name = accessDataTable.Columns[i].Caption;

                        //Create a MergeField name from the column name of the database table
                        string temp_bookmark_from_db_table = temp_column_name.Replace(" ", "_");
                        //MessageBox.Show(temp_bookmark_from_db_table);
                        
                        // **** FIELD REPLACEMENT IMPLEMENTATION GOES HERE ****//
                        if (temp_bookmark_from_db_table == fieldName || temp_bookmark_from_db_table == fieldName.Substring(0, fieldName.Length - 2))
                        {
                            string this_cell_data = this.accessDataSet.Tables[comboTables.Text].Rows[current_row_number][temp_column_name].ToString();
                            myMergeField.Select();
                            oWord.Selection.TypeText(this_cell_data);
                        }
                    }
                }
            }

            //oDoc.SaveAs("C:\\Law\\Documents\\WordTemplates\\myFile.doc");
            //oWord.Documents.Open("C:\\Law\\Documents\\WordTemplates\\myFile.doc");
            ////wordApp.Application.Quit();

            
            this.Close();
        }

       
    }
}
