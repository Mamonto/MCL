using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.XML.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Data.Common;

namespace DataEasy
{
    public partial class stripManus : Form
    {
        // Declare the ContextMenuStrip control.
        private ContextMenuStrip fruitContextMenuStrip;
        private ToolStripMenuItem fruitToolStripMenuItem;
        private MenuStrip ms;

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private Label label1;
        private TextBox textBox1;
        private Button button1;

        //Write SQL info to a TXT file
        private String sql;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(stripManus));
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 62);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Query Name";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(147, 59);
            this.textBox1.MaxLength = 40;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(211, 22);
            this.textBox1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(263, 120);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(95, 32);
            this.button1.TabIndex = 2;
            this.button1.Text = "Save";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // stripManus
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(398, 195);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(416, 240);
            this.Name = "stripManus";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Save Query to History List";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion


        //public stripManus(ContextMenuStrip context_menu_strip, MenuStrip menu_strip)
        public stripManus()
        {
            InitializeComponent();

            // Create a new ContextMenuStrip control.
            //fruitContextMenuStrip = new ContextMenuStrip();
            //fruitContextMenuStrip = context_menu_strip;

            // Attach an event handler for the 
            // ContextMenuStrip control's Opening event.
            //fruitContextMenuStrip.Opening += new System.ComponentModel.CancelEventHandler(cms_Opening);

            // Create a new MenuStrip control and add a ToolStripMenuItem.
            //ms = new MenuStrip();
            //fruitToolStripMenuItem = new ToolStripMenuItem("History", null, null, "History");
            //ms.Items.Add(fruitToolStripMenuItem);

            // Assign the MenuStrip control as the 
            // ToolStripMenuItem's DropDown menu.
            //fruitToolStripMenuItem.DropDown = fruitContextMenuStrip;

            // Add the MenuStrip control last.
            // This is important for correct placement in the z-order.
            // this.Controls.Add(ms);
        }

        //Write SQL info to a TXT file
        public void writeQueryToFile(string store_sql)
        { 
            sql = store_sql;

            this.Show();
            this.Focus();
        }

        //Save the SQL and provided name for it to the TXT sqlfile
        private void button1_Click(object sender, EventArgs e)
        {
            string messageError;
            //Write SQL info to a TXT file
            //string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = Application.StartupPath;
            string dbFile = filePath + @"\sqlfile.txt";
            string line;
            int count_lines = 0;
            int max_lines   = 10;

            //Calculate number of saved query records in a file
            //and later limit to a maximum allowed number
            try
            {
                System.IO.StreamReader reader = new System.IO.StreamReader(dbFile);
                while ((line = reader.ReadLine()) != null)
                {
                    //MessageBox.Show(line);
                    string[] items_1 = line.Split('\t');

                    if (items_1.Length == 2 && line != "")    //calculate number of lines that contain valid queries
                        count_lines += 1;
                }
                reader.Close();
            }
            catch (Exception) 
            {
                messageError = "Not able to read Query History file. It is deleted or moved from " + filePath
                                + ". A new history will be created.";
                MessageBox.Show(messageError, "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // Write the string to a file. Append new text to an existing file.
            // *********** Now, if the file is not found, the program throws an exception, but still opens a window
            // *********** for a user to enter a new query name and will save it to a newly created history file 'sqlfile.txt'
            try
            {
                if (this.textBox1.Text.Trim() == "" || this.textBox1.Text.Trim() == "\n")
                {
                    MessageBox.Show("Invalid name. Please re-enter.", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    //Write SQL to the History file
                    if (count_lines < max_lines)
                    {
                        System.IO.StreamWriter file = new System.IO.StreamWriter(dbFile, true);
                        file.WriteLine(this.textBox1.Text + "\t" + sql);
                        file.Close();
                    }
                    else
                    {
                        DialogResult r = MessageBox.Show("Number of saved queries cannot exceed " + max_lines + ". Would you like to replace the first query in the list?", "Exceeded Maximum Number of Queries  ", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                        if (r == DialogResult.Yes)
                        {

                            //*** delete this line first and then add a new one
                            string[] lines = File.ReadAllLines(dbFile);
                            List<string> list = new List<string>(lines);
                            list.RemoveAt(0);//remove item from index.
                            string[] newLines = list.ToArray();
                            File.WriteAllLines(dbFile, newLines);

                            System.IO.StreamWriter file = new System.IO.StreamWriter(dbFile, true);
                            file.WriteLine(this.textBox1.Text + "\t" + sql);
                            file.Close();
                        }
                    }
                    this.Close();
                }
            }
            catch (Exception)                   // *********** Verify EXCEPTION **************
            {
                messageError = "Not able to read Query History file. It is deleted or moved from " + filePath
                                + ".";
                MessageBox.Show(messageError, "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // This event handler is invoked when the ContextMenuStrip
        // control's Opening event is raised. It demonstrates
        // dynamic item addition and dynamic SourceControl 
        // determination with reuse.
        public void cms_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Acquire references to the owning control and item.
            Control c = fruitContextMenuStrip.SourceControl as Control;
            ToolStripDropDownItem tsi = fruitContextMenuStrip.OwnerItem as ToolStripDropDownItem;

            // Clear the ContextMenuStrip control's Items collection.
            fruitContextMenuStrip.Items.Clear();

            // Check the source control first.
            if (c != null)
            {
                // Add custom item (Form)
                fruitContextMenuStrip.Items.Add("Source: " + c.GetType().ToString());
            }
            else if (tsi != null)
            {
                // Add custom item (ToolStripDropDownButton or ToolStripMenuItem)
                fruitContextMenuStrip.Items.Add("Source: " + tsi.GetType().ToString());
            }

            //fruitContextMenuStrip.Items.Add("-");
            //fruitContextMenuStrip.Items.Add("Apples", null, this.dynamicMenu_Click);
            //fruitContextMenuStrip.Items.Add("Oranges", null, this.dynamicMenu_Click);
            //fruitContextMenuStrip.Items.Add("Pears", null, this.dynamicMenu_Click);

            // Set Cancel to false. 
            // It is optimized to true based on empty entry.
            e.Cancel = false;
        }

        //Read TXT file and find the SQL commands query that is stored in that file 
        private void dynamicMenu_Click(object sender, EventArgs e)
        {
            //string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = Application.StartupPath;
            string dbFile = filePath + @"\sqlfile.txt";
            string line;

            //This is the query name that was selected/clicked in the menu list by a user
            ToolStripMenuItem mi = (ToolStripMenuItem)sender;
            //MessageBox.Show("This is myMenuItemFile object. = " + mi.Text);

            // Open the text file using a stream reader.
            try
            {
                System.IO.StreamReader reader = new System.IO.StreamReader(dbFile);
                while ((line = reader.ReadLine()) != null)
                {
                    //MessageBox.Show(line);
                    string[] items = line.Split('\t');
                    if (items.Length == 2 && items[0] == mi.Text)
                        sql = items[1]; // Here's your sql query.
                }
                reader.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Error: This query was not found. It was deleted or the storage file was not located."); 
            }
            
            MessageBox.Show(mi.Text + "  " + sql);
        }

        //Return stored sql:
        public String getSql()
        {
            return this.sql;
        }

        //Return this object:
        public ContextMenuStrip getContextMenuStrip()
        {
            return this.fruitContextMenuStrip;
        }

        //Return this object:
        public ToolStripMenuItem getToolStripMenuItem()
        {
            return this.fruitToolStripMenuItem;
        }

        //Return MenuStrip object:
        public MenuStrip getMenuStripObj()
        {
            return this.ms;
        }
       
    }
}
