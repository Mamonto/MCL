//------------------------------------------------------------------------
//
// Author      : Mikhail Mamontov
// Date        : 23 April 2014
// Version     : 1.0
// Description : This Form Filters and 
//               Searches the database the user either has
//               to right click a column header in the Data
//               grid and chose from the menu or type a text search
//              
//
//------------------------------------------------------------------------

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DataEasy
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
    ///
	public class searchfrm : System.Windows.Forms.Form
	{
		#region "Private Declarations"
		private bool activateMouse = false; //a boolian to allow mouse click on the DGrid
		private object[][] elements;   //the filter menues' elements are stored here
		private string[] checkedMenu;  //the element you are looking for in the coulmn
		private DataSet accessDataSet = new DataSet();
		private OleDbConnection accessConnection = new  OleDbConnection();

		private string tableName="";  //DataBase table name
		private bool doUpdate = true; //Update the filter menues or not 
		private int columnHit;        //which column in the table is hit
		private System.Windows.Forms.DataGrid dGrid;
        private System.Windows.Forms.ContextMenu[] FilterMenu; //the filter menues for all the columns    //the combobox which holds the column names
        //to chose from in the search        //the search element is put here
        //for the text based search           //find row in data according to text based search        //remove all filters button
		private Form MFRM;
        private System.Windows.Forms.TextBox lblSelectString;

        // Excel object references.
        private Excel.Application m_objExcel = null;
        private Excel.Workbooks m_objBooks = null;
        private Excel._Workbook m_objBook = null;
        private Excel.Sheets m_objSheets = null;
        private Excel._Worksheet m_objSheet = null;
        private Excel.Range m_objRange = null;
        private Excel.Font m_objFont = null;
        private Excel.QueryTables m_objQryTables = null;
        private Excel._QueryTable m_objQryTable = null;
        // Frequenty-used variable for optional arguments.
        private object m_objOpt = System.Reflection.Missing.Value;
        // Paths used by the sample code for accessing and storing data.
        private object m_strSampleFolder = "C:\\TEMP\\ExcelData\\";
        private string m_strNorthwind = "C:\\TEMP\\Database63_be.mdb";
        private string QueryCommandText;
        private string CommandText = "";
        private TabControl TabControl;
        private TabPage tabPage1;
        private GroupBox groupBox2;
        private Label label2;
        private Button button1;
        private TextBox searchTxt;
        private ComboBox cBoxParamets;
        private Button Findbtn;
        private GroupBox groupBox1;
        private CheckBox checkBox1;
        private DateTimePicker dateTimePicker2;
        private DateTimePicker dateTimePicker1;
        private Label label1;
        private TextBox textBox1;
        private Button btnRestore;
        private TabPage tabPage2;
        private GroupBox groupBox3;
        private Button button2;
        private Label label3;
        private ComboBox cBoxParameters2; //query search

        //IncomingCallsFrm incomingcalls;
        private MainFrm mainform = null;

        //Dynamic Menues to Open Saved queries
        private Button btnSaveToHistory;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem historyToolStripMenuItem;
        private ToolStripMenuItem fruitToolStripMenuItem;
        private ContextMenuStrip fruitContextMenuStrip;
        private String sql;
        private ToolStripMenuItem closeToolStripMenuItem;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
        #endregion

         //The search class constructor
		public searchfrm(Form motherFrm, string datasource, string tablename, string SelectString)
		{
			InitializeComponent();

			//This refers to the parent form
			MFRM = motherFrm;
			//the tablename to do the search on
			tableName = tablename; 

			//Initializing the connection here to the source mdb file
			((System.ComponentModel.ISupportInitialize)(this.accessDataSet)).BeginInit();
			//accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + datasource;
            //accessConnection.Open();
            //
            //accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=Database63_be.mdb";
            string executable = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string path = (System.IO.Path.GetDirectoryName(executable));
            AppDomain.CurrentDomain.SetData("DataDirectory", path);
            
            this.accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=|DataDirectory|\Database63_fe.mdb";
            //accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=C:\\Users\\Mikhail\\Documents\\NEW MCL\\Database63_be.mdb";

            //m_strNorthwind = path;
            //this.accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source= " + m_strNorthwind + "\\Database63_fe.mdb";
            try
            {
                accessConnection.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Connecting to data source from this system... Please click OK", "Warning");
            }

            QueryCommandText = SelectString;

			loadData(SelectString);
			lblSelectString.Text = "Select Command = " + SelectString;
			activateMouse = true;
			accessConnection.Close();
			
			DataTable myDataTable = accessDataSet.Tables[tableName];

			//Find all columns and put them in the combobox
			//cBoxParamets
			for (int i=0;i<myDataTable.Columns.Count;i++)
			{
				cBoxParamets.Items.Add(myDataTable.Columns[i].Caption);
				if (i==0) cBoxParamets.Text = myDataTable.Columns[i].Caption;

                cBoxParameters2.Items.Add(myDataTable.Columns[i].Caption);
                if (i == 0) cBoxParameters2.Text = myDataTable.Columns[i].Caption;
			}
			dGrid.Height = this.Height-135;

            createDynamicMenus();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(searchfrm));
            this.dGrid = new System.Windows.Forms.DataGrid();
            this.lblSelectString = new System.Windows.Forms.TextBox();
            this.TabControl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btnSaveToHistory = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.searchTxt = new System.Windows.Forms.TextBox();
            this.cBoxParamets = new System.Windows.Forms.ComboBox();
            this.Findbtn = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnRestore = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cBoxParameters2 = new System.Windows.Forms.ComboBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.historyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dGrid)).BeginInit();
            this.TabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dGrid
            // 
            this.dGrid.AccessibleRole = System.Windows.Forms.AccessibleRole.ColumnHeader;
            this.dGrid.DataMember = "";
            this.dGrid.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dGrid.Location = new System.Drawing.Point(0, 283);
            this.dGrid.MaximumSize = new System.Drawing.Size(0, 465);
            this.dGrid.MinimumSize = new System.Drawing.Size(1059, 152);
            this.dGrid.Name = "dGrid";
            this.dGrid.ReadOnly = true;
            this.dGrid.Size = new System.Drawing.Size(1185, 152);
            this.dGrid.TabIndex = 1;
            this.dGrid.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dGrid_MouseDown);
            // 
            // lblSelectString
            // 
            this.lblSelectString.BackColor = System.Drawing.Color.Black;
            this.lblSelectString.ForeColor = System.Drawing.Color.LimeGreen;
            this.lblSelectString.Location = new System.Drawing.Point(16, 80);
            this.lblSelectString.Name = "lblSelectString";
            this.lblSelectString.ReadOnly = true;
            this.lblSelectString.Size = new System.Drawing.Size(312, 22);
            this.lblSelectString.TabIndex = 6;
            this.lblSelectString.Text = "Select Command";
            // 
            // TabControl
            // 
            this.TabControl.Controls.Add(this.tabPage1);
            this.TabControl.Controls.Add(this.tabPage2);
            this.TabControl.Dock = System.Windows.Forms.DockStyle.Top;
            this.TabControl.Location = new System.Drawing.Point(0, 28);
            this.TabControl.MinimumSize = new System.Drawing.Size(1062, 255);
            this.TabControl.Name = "TabControl";
            this.TabControl.SelectedIndex = 0;
            this.TabControl.Size = new System.Drawing.Size(1185, 255);
            this.TabControl.TabIndex = 15;
            // 
            // tabPage1
            // 
            this.tabPage1.AllowDrop = true;
            this.tabPage1.AutoScroll = true;
            this.tabPage1.BackColor = System.Drawing.SystemColors.ControlDark;
            this.tabPage1.Controls.Add(this.btnSaveToHistory);
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.btnRestore);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1177, 226);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Query";
            // 
            // btnSaveToHistory
            // 
            this.btnSaveToHistory.BackColor = System.Drawing.Color.YellowGreen;
            this.btnSaveToHistory.Location = new System.Drawing.Point(628, 144);
            this.btnSaveToHistory.Name = "btnSaveToHistory";
            this.btnSaveToHistory.Size = new System.Drawing.Size(79, 44);
            this.btnSaveToHistory.TabIndex = 19;
            this.btnSaveToHistory.Text = "Save to History";
            this.btnSaveToHistory.UseVisualStyleBackColor = false;
            this.btnSaveToHistory.Click += new System.EventHandler(this.saveSqlToHistoryFileTxt);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.searchTxt);
            this.groupBox2.Controls.Add(this.cBoxParamets);
            this.groupBox2.Controls.Add(this.Findbtn);
            this.groupBox2.Location = new System.Drawing.Point(363, 28);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(244, 171);
            this.groupBox2.TabIndex = 18;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Query";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(196, 17);
            this.label2.TabIndex = 4;
            this.label2.Text = "Select Field and Specific Data";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.button1.Location = new System.Drawing.Point(108, 103);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(79, 44);
            this.button1.TabIndex = 14;
            this.button1.Text = "Export to Excel";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // searchTxt
            // 
            this.searchTxt.Location = new System.Drawing.Point(20, 71);
            this.searchTxt.Name = "searchTxt";
            this.searchTxt.Size = new System.Drawing.Size(190, 22);
            this.searchTxt.TabIndex = 3;
            // 
            // cBoxParamets
            // 
            this.cBoxParamets.Location = new System.Drawing.Point(20, 40);
            this.cBoxParamets.Name = "cBoxParamets";
            this.cBoxParamets.Size = new System.Drawing.Size(190, 24);
            this.cBoxParamets.TabIndex = 2;
            this.cBoxParamets.Text = "Tables";
            // 
            // Findbtn
            // 
            this.Findbtn.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.Findbtn.Location = new System.Drawing.Point(20, 103);
            this.Findbtn.Name = "Findbtn";
            this.Findbtn.Size = new System.Drawing.Size(79, 44);
            this.Findbtn.TabIndex = 4;
            this.Findbtn.Text = "Query Results";
            this.Findbtn.UseVisualStyleBackColor = false;
            this.Findbtn.Click += new System.EventHandler(this.Findbtn_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Controls.Add(this.dateTimePicker2);
            this.groupBox1.Controls.Add(this.dateTimePicker1);
            this.groupBox1.Location = new System.Drawing.Point(71, 28);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(264, 101);
            this.groupBox1.TabIndex = 17;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Date Opened field";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(27, 17);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(233, 21);
            this.checkBox1.TabIndex = 12;
            this.checkBox1.Text = "Include range of dates in search";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.CustomFormat = "dd/MM/yyyy";
            this.dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker2.Location = new System.Drawing.Point(28, 65);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(169, 22);
            this.dateTimePicker2.TabIndex = 11;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.CustomFormat = "dd/MM/yyyy";
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker1.Location = new System.Drawing.Point(28, 35);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(169, 22);
            this.dateTimePicker1.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(76, 153);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(107, 17);
            this.label1.TabIndex = 16;
            this.label1.Text = "No. of Records:";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(189, 153);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(79, 22);
            this.textBox1.TabIndex = 15;
            // 
            // btnRestore
            // 
            this.btnRestore.BackColor = System.Drawing.Color.YellowGreen;
            this.btnRestore.Location = new System.Drawing.Point(278, 144);
            this.btnRestore.Name = "btnRestore";
            this.btnRestore.Size = new System.Drawing.Size(79, 44);
            this.btnRestore.TabIndex = 14;
            this.btnRestore.Text = "Restore";
            this.btnRestore.UseVisualStyleBackColor = false;
            this.btnRestore.Click += new System.EventHandler(this.btnRestore_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.SystemColors.ControlDark;
            this.tabPage2.Controls.Add(this.groupBox3);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1177, 226);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Stats";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button2);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.cBoxParameters2);
            this.groupBox3.Location = new System.Drawing.Point(71, 28);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(244, 171);
            this.groupBox3.TabIndex = 15;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Statustics";
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.button2.Location = new System.Drawing.Point(19, 103);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(79, 44);
            this.button2.TabIndex = 5;
            this.button2.Text = "Plot Chart";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.Chartbtn_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(81, 17);
            this.label3.TabIndex = 17;
            this.label3.Text = "Select Field";
            // 
            // cBoxParameters2
            // 
            this.cBoxParameters2.Location = new System.Drawing.Point(19, 40);
            this.cBoxParameters2.Name = "cBoxParameters2";
            this.cBoxParameters2.Size = new System.Drawing.Size(190, 24);
            this.cBoxParameters2.TabIndex = 15;
            this.cBoxParameters2.Text = "Tables";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.historyToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1185, 28);
            this.menuStrip1.TabIndex = 16;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // historyToolStripMenuItem
            // 
            this.historyToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.closeToolStripMenuItem});
            this.historyToolStripMenuItem.Name = "historyToolStripMenuItem";
            this.historyToolStripMenuItem.Size = new System.Drawing.Size(44, 24);
            this.historyToolStripMenuItem.Text = "File";
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(114, 24);
            this.closeToolStripMenuItem.Text = "Close";
            this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
            // 
            // searchfrm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.ClientSize = new System.Drawing.Size(1206, 362);
            this.Controls.Add(this.TabControl);
            this.Controls.Add(this.dGrid);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(1098, 200);
            this.Name = "searchfrm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MasterClient: Search DataBase";
            this.Resize += new System.EventHandler(this.searchfrm_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dGrid)).EndInit();
            this.TabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

        #region "Query History"
        //Dynamic Menues to Open Saved queries
        private void createDynamicMenus()
        {
            // Create a new ContextMenuStrip control.
            fruitContextMenuStrip = new ContextMenuStrip();
            //fruitContextMenuStrip = context_menu_strip;

            // Attach an event handler for the 
            // ContextMenuStrip control's Opening event.
            fruitContextMenuStrip.Opening += new System.ComponentModel.CancelEventHandler(cms_Opening);

            //MenuStrip ms = menu_strip;
            fruitToolStripMenuItem = new ToolStripMenuItem("History", null, null, "History");
            //menuStrip1.Items.Add(fruitToolStripMenuItem);

            // Assign the MenuStrip control as the 
            // ToolStripMenuItem's DropDown menu.
            fruitToolStripMenuItem.DropDown = fruitContextMenuStrip;

            menuStrip1.Items.Add(fruitToolStripMenuItem);
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
                //fruitContextMenuStrip.Items.Add("Source: " + c.GetType().ToString());
                fruitContextMenuStrip.Items.Add("Saved Queries:");
            }
            else if (tsi != null)
            {
                // Add custom item (ToolStripDropDownButton or ToolStripMenuItem)
                //fruitContextMenuStrip.Items.Add("Source: " + tsi.GetType().ToString());
                fruitContextMenuStrip.Items.Add("Saved Queries:");
            }

            // Populate the ContextMenuStrip control with its default items.
            // **************** Write this in a loop that reads data from the file and ADDs items as item[0]  *******************
            //fruitContextMenuStrip.Items.Add("-");
            //fruitContextMenuStrip.Items.Add("Apples", null, this.dynamicMenu_Click);
            //fruitContextMenuStrip.Items.Add("Oranges", null, this.dynamicMenu_Click);
            //fruitContextMenuStrip.Items.Add("Pears", null, this.dynamicMenu_Click);
            
            //string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = Application.StartupPath;
            string dbFile = filePath + @"\sqlfile.txt";
            string line;

            // Populate the ContextMenuStrip control with its default items.
            try
            {
                System.IO.StreamReader file = new System.IO.StreamReader(dbFile, true);
                fruitContextMenuStrip.Items.Add("-");
                while ((line = file.ReadLine()) != null)
                {
                    //MessageBox.Show(line);
                    string[] items = line.Split('\t');
                    if (items.Length == 2)
                        fruitContextMenuStrip.Items.Add(items[0], null, this.dynamicMenu_Click);
                }
                file.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Error: This query was not found. It was deleted or the storage file was not located.");
            }

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

            Cursor.Current = Cursors.WaitCursor;

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
                        QueryCommandText = items[1]; // Here's your sql query.
                }
                reader.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Error: This query was not found. It was deleted or the storage file was not located.");
            }

            MessageBox.Show(mi.Text + "  " + QueryCommandText);


            //searchfrm load_query = new searchfrm();
            loadData(QueryCommandText);
            Cursor.Current = Cursors.Default;
        }

        //Write Query to the TXT file
        private void saveSqlToHistoryFileTxt(object sender, System.EventArgs e)
        {
            //string filePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = Application.StartupPath;
            string dbFile = filePath + @"\sqlfile.txt";
            string thisQuery = QueryCommandText.Trim();
            string messageError;
            string line;

            // Open the text file using a stream reader.
            try
            {
                //Check for duplicate queries in the file of saved ones
                System.IO.StreamReader reader = new System.IO.StreamReader(dbFile);
                while ((line = reader.ReadLine()) != null)
                {
                    //MessageBox.Show(line);
                    string[] items_2 = line.Split('\t');

                    if (items_2.Length == 2 && items_2[1] == thisQuery)
                    {          
                        // ***************    Without break, This message will be displayed as many times as many 
                        //                      identical queries will be found in the file          ****************
                        DialogResult r = MessageBox.Show("This query already exist? Would you like to save the same query under a different name", "This Query Exists in History", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                        if (r == DialogResult.No)
                            thisQuery = ""; // Do not overwite existing query in the file.
                        else
                            break;
                    }
                }
                reader.Close();      
            }
            catch (Exception)  // *********** Verify EXCEPTION **************
            {
                messageError = "Not able to read Query History file. It is deleted or moved from " + filePath
                                + ". A new history will be created.";
                MessageBox.Show(messageError, "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
     

            //Save query to the file
            if (thisQuery != "")
            {
                mainform = new MainFrm();
                stripManus myStoreSql = new stripManus();
                
                if (!mainform.CheckOpened("Save Query to History List"))  //if this form has been already open, close it first
                    myStoreSql.writeQueryToFile(thisQuery);

                mainform = null;
                myStoreSql = null;
            }
        }
        #endregion

        #region "load Data"
		//This routine loads data from a mdb file
		//into the DGrid
		public void loadData(string SelectString)
		{
            //------------------------------------------------------------------------
            // Display today's date and time:
            //MessageBox.Show("Selected date is: " + DateTime.Today);

            //------------------------------------------------------------------------
			accessDataSet.RejectChanges();
			accessDataSet.Clear();

			OleDbCommand accessSelectCommand = new OleDbCommand();
			OleDbCommand accessInsertCommand = new OleDbCommand();
			OleDbDataAdapter accessDataAdapter = new OleDbDataAdapter();

			accessSelectCommand.CommandText = SelectString;
			accessSelectCommand.Connection =  accessConnection;
			accessDataAdapter.SelectCommand = accessSelectCommand;
				
			// Attempt to fill the dataset through the OleDbDataAdapter1.
			accessDataAdapter.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
																								   new System.Data.Common.DataTableMapping("Table", tableName)});				
			accessDataAdapter.Fill(accessDataSet);
				    
			dGrid.SetDataBinding(accessDataSet,tableName);

			int col = (accessDataSet.Tables[tableName].Columns.Count);
			int row = (accessDataSet.Tables[tableName].Rows.Count);
            
			if (doUpdate==true) checkedMenu = new String[col];
			elements = new object[col][];
			FilterMenu = new ContextMenu[col];

			for (int i = 0; i < col; i++) 
			{
				elements[i] = new object[row];
				if (doUpdate==true) checkedMenu[i] = "None";
			}

			for (int i=0; i < col; i++)
			{
				for (int j=0; j < row; j++)
				{ 
					if ((dGrid[j,i].GetType().Name == "Int32") ||
						(dGrid[j,i].GetType().Name == "DateTime")||
						(dGrid[j,i].GetType().Name == "Decimal"))
					{
						elements[i][j] = dGrid[j,i];
					}
					else elements[i][j] = dGrid[j,i].ToString();
				}
			}

			for (int i = 0; i < col; i++) 
			{
				try{Array.Sort(elements[i]);}
				catch
				{
					int newDim = 0;
					object[] tempArray = new object[elements[i].Length];
					for (int r = 0; r < elements[i].Length; r++)
					{
						if (elements[i][r].ToString()!= "")
						{
							tempArray[newDim] = elements[i][r];
							newDim++;
						}
					}
					elements[i] = new object[newDim];
					Array.Copy(tempArray,0,elements[i],0,newDim);
				}
				FilterMenu[i] = new ContextMenu();
				make_menues(elements[i],FilterMenu[i]);
			}
            int rows = accessDataSet.Tables[tableName].Rows.Count;
            textBox1.Text = rows.ToString();
		}
		#endregion

       
		
		#region "Make Menues"
		//This routine makes and updates filter menus
		//inaccordance with the displayed data in the grid
		private void make_menues(object[] array_elements, System.Windows.Forms.ContextMenu cMenu)
		{
			string Prev_Element = "";

			System.Windows.Forms.MenuItem mfirstItems = new MenuItem("None");
			mfirstItems.Click +=  new System.EventHandler(this.cMenuClick);
			cMenu.MenuItems.Add(mfirstItems);

			System.Windows.Forms.MenuItem[] mItems = new MenuItem[array_elements.Length];
			for (int i = 0; i < array_elements.Length;i++)
			{
				if (Prev_Element!=array_elements[i].ToString())
				{
					mItems[i] = new MenuItem(array_elements[i].ToString());
					mItems[i].Click +=  new System.EventHandler(this.cMenuClick);
					cMenu.MenuItems.Add(mItems[i]);
					Prev_Element = array_elements[i].ToString();
				}
			}
		}
		#endregion

		#region "Filter Menu Click Event"
		//This routine handles the filter menu click event
		private void cMenuClick(object sender, System.EventArgs e)
		{
			doUpdate = false;
			MenuItem tempItem = (MenuItem)sender;
			DataTable accessDataTable = accessDataSet.Tables[tableName];

			if ((accessDataTable.Columns[columnHit].DataType.ToString() == "System.Byte[]"))
			{
				MessageBox.Show("This DataType Cannot Be Filtered", "Unable To Do Filter", MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
				return;
			}

			checkedMenu[columnHit] = tempItem.Text;
			try
			{
				loadData(MakeSelectString(checkedMenu));
				lblSelectString.Text = "Select Command = " + MakeSelectString(checkedMenu); 
			}
			catch
			{
				MessageBox.Show("This DataType Cannot Be Filtered", "Unable To Do Filter", MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
				checkedMenu[columnHit] = "None";
				loadData(MakeSelectString(checkedMenu));
				lblSelectString.Text = "Select Command = " + MakeSelectString(checkedMenu); 
			}
		}
		#endregion

		#region "Make the Select Command"
		//This routine creates the seacrh command to be used
		//as a select command based on the options specified
		//by the user through text based search or menu based
		//filter
		private string MakeSelectString(string[] MenuChecked)
		{
			DataTable accessDataTable = accessDataSet.Tables[tableName];
			string STselect = "Select * From " + tableName + " Where ";
			bool there_is_Change = false;
			for (int i=0; i<MenuChecked.Length;i++)
			{
				
				string colType = accessDataTable.Columns[i].DataType.ToString();
				

				if (MenuChecked[i]!="None") 
				{
					if ((i!= 0) && (there_is_Change == true))
					{
						
						if (colType=="System.String")
						{
							STselect += " And [" + accessDataSet.Tables[tableName].Columns[i].Caption + " ] = '" + MenuChecked[i] + "'";
						}
						else if (colType=="System.DateTime")
						{
							STselect += " And [" + accessDataSet.Tables[tableName].Columns[i].Caption + " ] = #" + MenuChecked[i] + "#";
						}
						else	
						{
							STselect += " And [" + accessDataSet.Tables[tableName].Columns[i].Caption + " ] = " + MenuChecked[i];
						}
					}
					else 
					{
						if (colType=="System.String")
						{
							STselect += " [" + accessDataSet.Tables[tableName].Columns[i].Caption + " ] = '" + MenuChecked[i] + "'";
						}
						else if (colType=="System.DateTime")
						{
							STselect += " [" + accessDataSet.Tables[tableName].Columns[i].Caption + " ] = #" + MenuChecked[i] + "#";
						}
						else
						{
							STselect += " [" + accessDataSet.Tables[tableName].Columns[i].Caption + " ] = " + MenuChecked[i];
						}
					}
					there_is_Change = true;
				}
			}
			if (there_is_Change == false) STselect = "Select * From " + tableName;
			lblSelectString.Text = "Select Command = " + STselect; 
			return STselect;
		}
		#endregion

		#region "Data Grid Mouse Down Event"
		//This routine creates and loads filter menues for 
		//the datagrid and then displays them if the user right
		//clicks the header of any column
		private void dGrid_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (activateMouse == false) return;
			if (e.Button != System.Windows.Forms.MouseButtons.Right) return;
			DataGrid myGrid = (DataGrid) sender;
			System.Windows.Forms.DataGrid.HitTestInfo hti;
			hti = myGrid.HitTest(e.X, e.Y);
			string message = "You clicked ";

			switch (hti.Type) 
			{
				case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader :
					message += "the column header for column " + hti.Column;
					columnHit = hti.Column;
					FilterMenu[hti.Column].Show(dGrid,new Point(e.X,e.Y));
					break;
			}
		}
		#endregion

		#region "Text based search"
		//Button find is clicked
		private void Findbtn_Click(object sender, System.EventArgs e)
		{
            //buttonFind_Click(sender, e);
            find_the_data();
            //find_the_data_2();
		}

        private void Chartbtn_Click(object sender, EventArgs e)
        {
            plot_chart();
        }

        /*
        //Search a range of dates. Create select dates from DateTimePicker
        public static DataSet SearchDateRange(DateTime firstdate, DateTime seconddate)
        {
            SqlConnection conn = new SqlConnection("Data source = blah; initial catalog=yipps; integrated security=true");
            DataSet dst = new DataSet();
            SqlDataAdapter dtr = new SqlDataAdapter("SELECT * FROM dbo.ImportedCashins WHERE date >= @StartDate AND date < @EndDate", conn);
            dtr.SelectCommand.Parameters.AddWithValue("@StartDate", firstdate);
            dtr.SelectCommand.Parameters.AddWithValue("@EndDate", seconddate);
            dtr.Fill(dst);
            return dst;
        }
        */

		

        private void find_the_data()
        {
            //... BETWEEN #" +FromDateTimePicker.Value.ToLongDateString() + "# AND ...
            //dateTimePicker1.Value.Date    
            //+ cBoxParamets.Text +

            int index = 0;
            if ((searchTxt.Text == "") && (checkBox1.Checked == false) ) return;
            DataTable accessDataTable = accessDataSet.Tables[tableName];

            
            string CommandText_1    = "SELECT * FROM " + tableName + " Where ([" + cBoxParamets.Text + "] = ";
            string CommandText_3 = " AND ([" + cBoxParamets.Text + "] = ";

            //string CommandText_3    = " OR [" + cBoxParamets.Text + "] =  " + " '" + searchTxt.Text + "')";
            //string CommandText_2 = " AND ([Date Opened] BETWEEN #" + dateTimePicker1.Value.ToShortDateString() + "# AND #" + dateTimePicker2.Value.ToShortDateString() + "#)";           
            //string CommandText_2 = " AND [Date Opened] BETWEEN #01/07/2013# AND #31/07/2013#"; 
            //string CommandText = "SELECT * FROM " + tableName + " Where [" + cBoxParamets.Text + "] = 'Shelley Levine' AND [Hearing Date] BETWEEN #" + dateTimePicker1.Value.Date + "# AND #" + dateTimePicker2.Value.Date + "#";

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // NOTE: there was an issue with searching the DateTime in the Database: 
            // The programe would select dateTimePicker from the calendar in the correct Type=Short format (dd/MM/yyyy) and it will display 
            // in the dGrid table the searched dates found in the database also format (dd/MM/yyyy).
            // However, when the program is searching the database it automatically reverses the dateTimePicker to (MM/dd/yyyy).
            // For example, if search is to look for 01/07/2014 (1-July-2014), the internal SQL search will look for 07/01/2014 in the database
            // still formatted Date Type=(dd/MM/yyyy). The search will result in 7-January-2014 with dGrid displaying 07/01/2014.
            // Solution patch: 
            //      1) get the dateTimePicker in the format dd/MM/yyyy. Save as DataType
            //      2) reverse date and month by changing the dateTimePicker variable Format to MM/dd/yyyy. Save it to a String variable
            //      3) search the SQL Database pretending dd->MM and MM->dd
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            DateTime dt1 = dateTimePicker1.Value.Date;
            DateTime dt2 = dateTimePicker2.Value.Date;
            //string dt1_str = Convert.ToDateTime(dt1.ToString("MM/dd/yyyy"));
            //string dt2_str = Convert.ToDateTime(dt2.ToString("MM/dd/yyyy"));
            //String.Format("{0:M/d/yyyy}", dt1);            // "3/9/2008"
            //String.Format("{0:M/d/yyyy}", dt2);            // "3/9/2008"

            //CommandText = "SELECT * FROM " + tableName + " Where [Date Opened] BETWEEN #7/1/2013# AND #7/31/2013#";
            //CommandText = "SELECT * FROM " + tableName + " Where (([" + cBoxParamets.Text + "] = 'Shelley Levine') AND ([Date Opened] = #17/06/1905#))";
            //CommandText = "SELECT * FROM " + tableName + " Where [Date Opened] BETWEEN #" + dateTimePicker1.Value.ToShortDateString() + "# AND #" + dateTimePicker2.Value.ToShortDateString() + "#";
            //string CommandText_2 = "SELECT * FROM " + tableName + " Where [Date Opened] BETWEEN #" + String.Format("{0:M/d/yyyy}", dt1) +"# AND #" + String.Format("{0:M/d/yyyy}", dt2) +"#";
            string CommandText_2 = " AND [Date Opened] BETWEEN #" + String.Format("{0:M/d/yyyy}", dt1) + "# AND #" + String.Format("{0:M/d/yyyy}", dt2) + "#)";
            
                                                                                  
            try
            {               
                for (int i = 0; i < accessDataSet.Tables[tableName].Columns.Count; i++)
                {
                    if (cBoxParamets.Text == cBoxParamets.Items[i].ToString()) 
                    { 
                        index = i;
                        checkedMenu[index] = searchTxt.Text;
                    }
                    //checkedMenu[i] = "None";
                }

                //string CommandText = "";
                if (accessDataTable.Columns[index].DataType.ToString() == "System.Byte[]")
                {
                    MessageBox.Show("This DataType Cannot Be Filtered", "Unable To Do Filter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                
				if (accessDataTable.Columns[index].DataType.ToString()=="System.String")
				{
                    if (CommandText == "")
                        CommandText_1 += " '" + searchTxt.Text + "'";
                    else
                        CommandText_3 += " '" + searchTxt.Text + "'";
				}
				else if (accessDataTable.Columns[index].DataType.ToString()=="System.DateTime")
				{					
                    if (CommandText == "")
                        CommandText_1 += " #" + searchTxt.Text + "#";
                    else
                        CommandText_3 += " #" + searchTxt.Text + "#";
				}
				else
				{
					//CommandText += searchTxt.Text;
				}

                //Concatonate both strings in one search: range of dates and  selected data to look for.
                if (checkBox1.Checked == true && searchTxt.Text != "")
                {
                    if (searchTxt.Text != "all" && searchTxt.Text != "")
                    {
                        if (CommandText == "")
                            CommandText = CommandText_1 + CommandText_2;
                        else
                            CommandText = CommandText + CommandText_3 + CommandText_2;
                    }
                    else
                    {
                        if (CommandText == "")
                            CommandText = "SELECT * FROM " + tableName + " Where [Date Opened] BETWEEN #" + String.Format("{0:M/d/yyyy}", dt1) + "# AND #" + String.Format("{0:M/d/yyyy}", dt2) + "#";
                        else
                            return;  //nothing to do, because selected all on the existing query
                    }
                }
                else if (checkBox1.Checked == false && searchTxt.Text != "")
                {
                    if (searchTxt.Text != "all" && searchTxt.Text != "")
                    {
                        if (CommandText == "")
                            CommandText = CommandText_1 + ")";
                        else
                            CommandText = CommandText + CommandText_3 + ")";
                    }
                    else
                        return;
                        //CommandText = "SELECT * FROM " + tableName;
                }
                else if (checkBox1.Checked == true && searchTxt.Text == "")
                {
                    if (CommandText == "")
                        CommandText = CommandText_1 + ")";
                    else
                        CommandText = CommandText + CommandText_3 + ")";
                }
                
                QueryCommandText = CommandText;
                loadData(CommandText);
                //lblSelectString.Text = "Select * From " + CommandText;

                //Save the CommandText variable 
                QueryCommandText = CommandText;

                ////////////////////////////
                // Start a new workbook in Excel.
                /*
                m_objExcel = new Excel.Application();
                m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
                m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));

                // Create a QueryTable that starts at cell A1.
                m_objSheets = (Excel.Sheets)m_objBook.Worksheets;
                m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item(1));
                m_objRange = m_objSheet.get_Range("A1", m_objOpt);
                m_objQryTables = m_objSheet.QueryTables;
                m_objQryTable = (Excel._QueryTable)m_objQryTables.Add(
                    "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                    m_strNorthwind + ";", m_objRange, CommandText);   // "Select * From " + tableName
                m_objQryTable.RefreshStyle = Excel.XlCellInsertionMode.xlInsertEntireRows;
                m_objQryTable.Refresh(false);

                
                // Save the workbook and quit Excel.
                m_objBook.SaveAs(m_strSampleFolder + "InsertDataTable.xls", m_objOpt, m_objOpt,
                    m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlNoChange, m_objOpt, m_objOpt,
                    m_objOpt, m_objOpt, m_objOpt);
                m_objBook.Close(false, m_objOpt, m_objOpt);
                m_objExcel.Quit();
                */
                
                ////////////////////////////////////////////////////////////////////////
            }
            catch
            {
                MessageBox.Show("This search string you specified does not match the datatype of the column!!" +
                    " OR There is no data in the Table", "Unable To Do Filter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                loadData("Select * From " + tableName);
                //lblSelectString.Text = "Select Command = Select * From " + tableName;
                searchTxt.Text = "";
            }
        }

        //Plot Chart in Excel
        private void plot_chart()
        {
            //20150425
            String searchValue = this.cBoxParameters2.Text;

            int rowIndex = -1;

            dGrid.UnSelect(dGrid.CurrentRowIndex);
            Cursor.Current = Cursors.WaitCursor;

            int i = this.BindingContext[accessDataSet, tableName].Position;
            int number_of_rows_in_table = this.accessDataSet.Tables[tableName].Rows.Count;
            int number_of_clmns_in_table = this.accessDataSet.Tables[tableName].Columns.Count;

            int index_column = this.accessDataSet.Tables[tableName].Columns[searchValue].Ordinal;
            string headerText = this.accessDataSet.Tables[tableName].Columns[index_column].ColumnName.ToString();
            string[] array = new string[number_of_rows_in_table];

            for (int var_count_rows_in_loop = 0; var_count_rows_in_loop < number_of_rows_in_table; var_count_rows_in_loop++)
            {
                if (dGrid[var_count_rows_in_loop, index_column].ToString() != "")
                {
                    array[var_count_rows_in_loop] = dGrid[var_count_rows_in_loop, index_column].ToString();
                }
            }
            
            Dictionary<string, int> dict = new Dictionary<string, int>();
            try
            {
                    foreach (string value in array)
                    {
                        if (value != null)
                        {
                            if (dict.ContainsKey(value))
                                dict[value]++;
                            else
                                dict[value] = 1;
                        }
                    }
            }
            catch
            {
                MessageBox.Show("Error message", "Unable To Do Filter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            foreach (KeyValuePair<string, int> kv in dict)
            {
                //MessageBox.Show(kv.Key.ToString());   //name
                //MessageBox.Show(kv.Value.ToString()); //count
            }
            /*
            Random rdn = new Random();
            for (int i = 0; i < 50; i++)
            {
                chart1.Series["test1"].Points.AddXY
                                (rdn.Next(0,10), rdn.Next(0,10));
                chart1.Series["test2"].Points.AddXY
                                (rdn.Next(0,10), rdn.Next(0,10));
            }
                
            chart1.Series["test1"].ChartType = 
                                SeriesChartType.FastLine;
            chart1.Series["test1"].Color = Color.Red;

            chart1.Series["test2"].ChartType = 
                                SeriesChartType.FastLine;
            chart1.Series["test2"].Color = Color.Blue; 
            */


            Form display_textField_in_window = new ChartStats(dict, headerText);
            try
            {
                display_textField_in_window.Show();
            }
            catch (Exception)
            {
                try
                {
                    MessageBox.Show("Warning: Attempt to create Stats Chart.");
                }
                catch (Exception) { }
            }
            //20150425  END  //////////////////////////////////////////////////////////////////////////
            /*
            m_objExcel = new Excel.Application();
            m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
            m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));

            // Create a QueryTable that starts at cell A1.
            m_objSheets = (Excel.Sheets)m_objBook.Worksheets;
            m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item(1));
            m_objRange = m_objSheet.get_Range("A1", m_objOpt);
            m_objQryTables = m_objSheet.QueryTables;

            string executable = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string path = (System.IO.Path.GetDirectoryName(executable));
           
            m_objQryTable = (Excel._QueryTable)m_objQryTables.Add("OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\Mikhail\\Documents\\NEW MCL\\Database63_be.mdb;", m_objRange, QueryCommandText);


            m_objQryTable.RefreshStyle = Excel.XlCellInsertionMode.xlInsertEntireRows;
            m_objQryTable.Refresh(false);


            ////////////////////////////////////////////////////////////////////////
            String sMsg;
            int iNumQtrs;
            Excel.Range oResizeRange;
            Excel._Chart oChart;
            Excel.Series oSeries;

            //Determine how many quarters to display data for.
            for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
            {
                sMsg = "Enter sales data for ";
                sMsg = String.Concat(sMsg, iNumQtrs);
                sMsg = String.Concat(sMsg, " quarter(s)?");
     
                DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
                    MessageBoxButtons.YesNo);
                if (iRet == DialogResult.Yes)
                    break;
            }

            sMsg = "Displaying data for ";
            sMsg = String.Concat(sMsg, iNumQtrs);
            sMsg = String.Concat(sMsg, " quarter(s).");

            //MessageBox.Show(sMsg, "Quarterly Sales");

            //Starting at E1, fill headers for the number of columns selected.
            //m_objRange      = m_objSheet.get_Range("A1", m_objOpt);
            oResizeRange = m_objSheet.get_Range("AR1", "AR1").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

            //Change the Orientation and WrapText properties for the headers.
            oResizeRange.Orientation = 38;
            oResizeRange.WrapText = true;

            //Fill the interior color of the headers.
            oResizeRange.Interior.ColorIndex = 36;

            //Fill the columns with a formula and apply a number format.
            oResizeRange = m_objSheet.get_Range("AR2", "AR6").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=RAND()*100";
            oResizeRange.NumberFormat = "$0.00";

            //Add a Totals formula for the sales data and apply a border.
            oResizeRange = m_objSheet.get_Range("AR8", "AR8").get_Resize(Missing.Value, iNumQtrs);
            oResizeRange.Formula = "=SUM(AR2:AR6)";

            //Add a Chart for the selected data.
            m_objBook = (Excel._Workbook)m_objSheet.Parent;
            oChart = (Excel._Chart)m_objBook.Charts.Add(Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);

            //Use the ChartWizard to create a new chart from the selected data.
            oResizeRange = m_objSheet.get_Range("AR2:AR6", Missing.Value).get_Resize(
                Missing.Value, iNumQtrs);
            oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value,
                Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            oSeries = (Excel.Series)oChart.SeriesCollection(1);
            oSeries.XValues = m_objSheet.get_Range("C2", "C6");
            for (int iRet = 1; iRet <= iNumQtrs; iRet++)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
                String seriesName;
                seriesName = "=\"Q";
                seriesName = String.Concat(seriesName, iRet);
                seriesName = String.Concat(seriesName, "\"");
                oSeries.Name = seriesName;
            }

            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, m_objSheet.Name);

            //Move the chart so as not to cover your data.
            oResizeRange = (Excel.Range)m_objSheet.Rows.get_Item(10, Missing.Value);
            m_objSheet.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
            oResizeRange = (Excel.Range)m_objSheet.Columns.get_Item(2, Missing.Value);
            m_objSheet.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;

            //Make sure Excel is visible and give the user control
            //of Microsoft Excel's lifetime.
            m_objExcel.Visible = true;
            m_objExcel.UserControl = true;

            //////////////////////////////////////////////////////////
            //m_objExcel.Quit();
            */
        }


		//This routine removes all the filters
		//and displays all the data
		private void btnRestore_Click(object sender, System.EventArgs e)
		{
			for (int i =0 ; i<accessDataSet.Tables[tableName].Columns.Count;i++)
			{
				checkedMenu[i] = "None";
			}
            QueryCommandText = "Select * From " + tableName;
            loadData(QueryCommandText);
			lblSelectString.Text = "Select Command = Select * From " + tableName;

            //reset query search
            CommandText = "";
		}

		//The user clicked enter instead of buttonFind
		//should give same affect
		private void searchTxt_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{find_the_data();} 
		}

		#endregion

		#region "search Form events"
		//From Resize routine
		private void searchfrm_Resize(object sender, System.EventArgs e)
		{
			dGrid.Height = this.Height-135;
		}

		//Form Closing Routine
		private void searchfrm_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			MFRM.Enabled = true;
		}
		#endregion

        private void button1_Click(object sender, EventArgs e)
        {
            int rows = accessDataSet.Tables[tableName].Rows.Count;
            //MessageBox.Show("No. of Rows are : " + rows.ToString());
            textBox1.Text = rows.ToString();
        }    
        
        private void buttonFind_Click(object sender, System.EventArgs e)
        {
            // Display the selected date and time:
            MessageBox.Show("Your've selected the meeting date: "
            + dateTimePicker1.Value.Date);

            // Display today's date and time:
            MessageBox.Show("Today is: " + DateTime.Today);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool checkBox1Checked = true;
        }

        ////////////////////////////////////
        // Save DataSet to Excel xml file
        private void button1_Click_1(object sender, System.EventArgs e)
        {
            m_objExcel = new Excel.Application();
            m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
            m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));

            // Create a QueryTable that starts at cell A1.
            m_objSheets = (Excel.Sheets)m_objBook.Worksheets;
            m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item(1));
            m_objRange = m_objSheet.get_Range("A1", m_objOpt);
            m_objQryTables = m_objSheet.QueryTables;

            string executable = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string path = (System.IO.Path.GetDirectoryName(executable));
 
            m_objQryTable = (Excel._QueryTable)m_objQryTables.Add(
                "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                path + "\\Database63_fe.mdb;", m_objRange, QueryCommandText);

            m_objQryTable.RefreshStyle = Excel.XlCellInsertionMode.xlInsertEntireRows;
            m_objQryTable.Refresh(false);

            //Make sure Excel is visible and give the user control
            //of Microsoft Excel's lifetime.
            m_objExcel.Visible = true;
            m_objExcel.UserControl = true;
        }


        public void button1_Click_1bkp(object sender, EventArgs e)
        {
            // Start a new workbook in Excel.
            m_objExcel = new Excel.Application();
            m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
            m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));

            // Create a QueryTable that starts at cell A1.
            m_objSheets = (Excel.Sheets)m_objBook.Worksheets;
            m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item(1));
            m_objRange = m_objSheet.get_Range("A1", m_objOpt);
            m_objQryTables = m_objSheet.QueryTables;
            m_objQryTable = (Excel._QueryTable)m_objQryTables.Add(
                "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                m_strNorthwind + ";", m_objRange, QueryCommandText);   // "Select * From " + tableName
            m_objQryTable.RefreshStyle = Excel.XlCellInsertionMode.xlInsertEntireRows;



            string col1 = "";
            //string table_no = type;
            col1 = "Lawyer Handling File";
            System.Data.DataRowCollection dr = accessDataSet.Tables[tableName].Rows;
            int cols = accessDataSet.Tables[tableName].Columns.Count;
            //ExcelControl1.Cells[1, 1] = col1;
            for (int i = 0; i < cols; i++)
            {
                col1 = accessDataSet.Tables[tableName].Columns[i].ColumnName;
                //ExcelControl1.Cells[2, i + 1] = col1;
            }

            int num = dr.Count;
            for (int i = 0; i < num; i++)
            {
                object[] array = dr[i].ItemArray;
                int j;
                for (j = 0; j < array.Length; j++)
                {
                    col1 = array[j].ToString();
                    //ExcelControl1.Cells[i + 3, j + 1] = col1;

                }

            }
            //ExcelControl1.Show();
        }

        //You can override OnFormClosing to do this. Just be careful you don't do 
        //anything too unexpected, as clicking the 'X' to close is a well understood behavior.
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            mainform = new MainFrm();
            base.OnFormClosing(e);

            if (e.CloseReason == CloseReason.WindowsShutDown
                || e.CloseReason == CloseReason.TaskManagerClosing) return;

            DialogResult r = MessageBox.Show("Are you sure you want to close Query?", "Close Query", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (r == DialogResult.Yes)
            {
                //this.Close();       
                if (MFRM.Enabled == false)
                {
                    this.Hide();
                    MFRM.Enabled = true;
                    MFRM.Focus();
                }
                else
                {
                    this.Hide();
                    mainform.StartFormOpen();
                }
            }
            else
               e.Cancel = true;

            mainform = null;
            Cursor.Current = Cursors.Arrow;
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            mainform = new MainFrm();

            DialogResult r = MessageBox.Show("Are you sure you want to close Query?", "Close Query", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (r == DialogResult.Yes)
            {
                //this.Close();       
                if (MFRM.Enabled == false)
                {
                    this.Hide();
                    MFRM.Enabled = true;
                    MFRM.Focus();
                }
                else
                {
                    this.Hide();
                    mainform.StartFormOpen();
                }
            }
            //else
            //    e.Cancel = true;

            mainform = null;
            Cursor.Current = Cursors.Arrow;
        }

        //####################################################################################################
        // UNUSED FUNCTIONS
        //####################################################################################################

        private void loadData_2(string SelectString)
        {
            accessDataSet.RejectChanges();
            accessDataSet.Clear();

            OleDbCommand accessSelectCommand = new OleDbCommand();
            OleDbCommand accessInsertCommand = new OleDbCommand();
            OleDbDataAdapter accessDataAdapter = new OleDbDataAdapter();

            accessSelectCommand.CommandText = SelectString;
            accessSelectCommand.Connection = accessConnection;
            accessDataAdapter.SelectCommand = accessSelectCommand;

            // Attempt to fill the dataset through the OleDbDataAdapter1.
            //accessDataAdapter.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            //																					   new System.Data.Common.DataTableMapping("Table", tableName)});
            accessDataAdapter.Fill(accessDataSet);

            dGrid.SetDataBinding(accessDataSet, tableName);

            int col = (accessDataSet.Tables[tableName].Columns.Count);
            int row = (accessDataSet.Tables[tableName].Rows.Count);

            if (doUpdate == true) checkedMenu = new String[col];
            elements = new object[col][];
            FilterMenu = new ContextMenu[col];

            for (int i = 0; i < col; i++)
            {
                elements[i] = new object[row];
                if (doUpdate == true) checkedMenu[i] = "None";
            }

            for (int i = 0; i < col; i++)
            {
                for (int j = 0; j < row; j++)
                {
                    if ((dGrid[j, i].GetType().Name == "Int32") ||
                        (dGrid[j, i].GetType().Name == "DateTime") ||
                        (dGrid[j, i].GetType().Name == "Decimal"))
                    {
                        elements[i][j] = dGrid[j, i];
                    }
                    else elements[i][j] = dGrid[j, i].ToString();
                }
            }

            for (int i = 0; i < col; i++)
            {
                try { Array.Sort(elements[i]); }
                catch
                {
                    int newDim = 0;
                    object[] tempArray = new object[elements[i].Length];
                    for (int r = 0; r < elements[i].Length; r++)
                    {
                        if (elements[i][r].ToString() != "")
                        {
                            tempArray[newDim] = elements[i][r];
                            newDim++;
                        }
                    }
                    elements[i] = new object[newDim];
                    Array.Copy(tempArray, 0, elements[i], 0, newDim);
                }
                FilterMenu[i] = new ContextMenu();
                make_menues(elements[i], FilterMenu[i]);
            }
            int rows = accessDataSet.Tables[tableName].Rows.Count;
            textBox1.Text = rows.ToString();
        }

        /////////////////////////////////////////////////////
        //based on the element required in the search string
        private void find_the_data_2()
        {
            int index = 0;
            //if (searchTxt.Text == "") return;
            DataTable accessDataTable = accessDataSet.Tables[tableName];
            string CommandText = "SELECT * FROM " + tableName + " Where [" + cBoxParamets.Text + "] = ";

            try
            {
                for (int i = 0; i < accessDataSet.Tables[tableName].Columns.Count; i++)
                {
                    if (cBoxParamets.Text == cBoxParamets.Items[i].ToString()) { index = i; }
                    checkedMenu[i] = "None";
                }

                if (accessDataTable.Columns[index].DataType.ToString() == "System.Byte[]")
                {
                    MessageBox.Show("This DataType Cannot Be Filtered", "Unable To Do Filter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (accessDataTable.Columns[index].DataType.ToString() == "System.String")
                {
                    CommandText += " '" + searchTxt.Text + "'";
                }
                else if (accessDataTable.Columns[index].DataType.ToString() == "System.DateTime")
                {
                    CommandText += " #" + searchTxt.Text + "#";
                }
                else
                {
                    CommandText += searchTxt.Text;
                }

                loadData(CommandText);
                lblSelectString.Text = "Select Command = " + CommandText;
            }
            catch
            {
                MessageBox.Show("This search string you specified does not match the datatype of the column!!!!!!");
                loadData("Select * From " + tableName);
                lblSelectString.Text = "Select Command = Select * From " + tableName;
                searchTxt.Text = "";
            }

        }

        //based on the element required in the search string
        private void find_the_data_3()
        {
            int index = 0;
            //if (searchTxt.Text == "") return;
            DataTable accessDataTable = accessDataSet.Tables[tableName];
            string CommandText = "SELECT * FROM " + tableName + " Where [" + cBoxParamets.Text + "] = ";
            string CommandText_2 = "SELECT * FROM " + tableName + " WHERE [Hearing Date] BETWEEN #" + dateTimePicker1.Value.Date + "# AND #" + dateTimePicker2.Value.Date + "#";

            try
            {
                for (int i = 0; i < accessDataSet.Tables[tableName].Columns.Count; i++)
                {
                    if (cBoxParamets.Text == cBoxParamets.Items[i].ToString()) { index = i; }
                    checkedMenu[i] = "None";
                }

                if (accessDataTable.Columns[index].DataType.ToString() == "System.Byte[]")
                {
                    MessageBox.Show("This DataType Cannot Be Filtered", "Unable To Do Filter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (accessDataTable.Columns[index].DataType.ToString() == "System.String")
                {
                    CommandText += " '" + searchTxt.Text + "'";
                    //CommandText_2 += " '" + searchTxt.Text + "'";
                }
                else if (accessDataTable.Columns[index].DataType.ToString() == "System.DateTime")
                {
                    //CommandText += " #" + searchTxt.Text + "#";
                    CommandText_2 += " #" + searchTxt.Text + "#";
                }
                else
                {
                    CommandText += searchTxt.Text;
                    CommandText_2 += searchTxt.Text;
                }

                loadData(CommandText);
                loadData_2(CommandText_2);
                lblSelectString.Text = "Select Command = " + CommandText;
            }
            catch
            {
                MessageBox.Show("This search string you specified does not match the datatype of the column!!!!!!");
                loadData("Select * From " + tableName);
                lblSelectString.Text = "Select Command = Select * From " + tableName;
                searchTxt.Text = "";
            }
        }

        
 
 
        //##################     UNUSED FUNCTIONS end     ##########################
    
	}
}
