//------------------------------------------------------------------------
//
// Author      : Mikhail Mamontov.
// Date        : 2015
// Version     : 1.0
// Description : An easy Application to link with an access mdb
//               file to allow the user to Delete (optional), Select
//               Update the dataloaded
//               Filter and Search capabilities are also
//               included
//               NiceMenu and CreateControls classes are done by 
//               other authors I downloaded them from planet-source-code
//               site I thank them for sharing their codes here in planet-source-code
//
//------------------------------------------------------------------------


using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Utility.NiceMenu;
using System.Data.OleDb;

using Microsoft.Office;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Web;
using System.Windows.Controls.Primitives;


namespace DataEasy
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
    public class IncomingCallsFrm : System.Windows.Forms.Form
	{
		#region "Private Declarations" 
        private string remeberLastKey = "";
        private string[] checkedMenu;  //the element you are looking for in the coulmn    //the combobox which holds the column names
        private object[][] elements;   //the filter menues' elements are stored here
        //private DataSet accessDataSet = new DataSet();
        private string tableName = "";  //DataBase table name
        public int count_all_rows_in_loop;

        private System.Windows.Forms.Button btnNavPrev;
		private System.Windows.Forms.Label lblNavLocation;
		private System.Windows.Forms.Button btnNavNext;
		private System.Windows.Forms.Button btnAdd;
		private System.Windows.Forms.Button btnDelete;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button Findbtn;
		private System.Windows.Forms.DataGrid dGrid;
		private System.Windows.Forms.ImageList imageMenu;
        private System.Windows.Forms.MainMenu menuFile;
        private System.Windows.Forms.MenuItem menuF_open;
		private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ComboBox comboTables;
        private System.Windows.Forms.TextBox textBox2;
		
		private string[] colType;         //This array holds all the columnTypes;
		private System.Windows.Forms.ContextMenu[] AutoMenu; //AutoDate or AutoNumber Options
		private int GPostion;             //DataGrid's postion from Top
		private bool DataLoaded = false;  //Check if the data is loaded into the system
		private NiceMenu myNiceMenu;      //the menues with icons
        private DataSet accessDataSet;    //the main DataSet
		private string ComboBoxText="";   //the item selected in the ComboTables control
		private string dataSourceFile=""; //source file of the .mdb file
		private OleDbDataAdapter accessDataAdapter; //the adapter to be used in conjunction with
        //the database file 
        private ToolStrip toolStrip1;
        private ToolStripButton toolStripButton1;
        private ToolStripButton toolStripButton2;
        private ToolStripButton toolStripButton3;
        private ToolStripButton toolStripButton4;
        private ToolStripButton toolStripButton5;
        private ToolStripButton toolStripButton6;
        private ToolStripButton toolStripButton7;
        private ToolStripLabel toolStripLabel1;
        private ToolStripButton toolStripButton8;
        private RichTextBox textBox1;
        //private Button btnWordDoc;
        private ToolStripButton toolStripButton9;
        
        //private ComboBox cBoxParamets;
        //private Button Findbtn;
        //private MenuItem menuItem3;
		private OleDbConnection accessConnection = new  OleDbConnection();
        private OleDbConnection store_accessConnection = new OleDbConnection();
        private RichTextBox richTextBox2;
        private RichTextBox richTextBox3;
        private RichTextBox richTextBox4;
        private RichTextBox richTextBox5;
        private Label label1;
        private Label label2;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem saveToolStripMenuItem;
        private ToolStripMenuItem closeToolStripMenuItem;
        private ToolStripMenuItem toolsToolStripMenuItem;
        private ToolStripMenuItem mailMergeToolStripMenuItem;
        //private DataView dv;
        private DataGridView dataGridView1 = new DataGridView();

        private Form MFRM;
        private MainFrm mainform = null;
        
		#endregion

        public IncomingCallsFrm(Form motherFrm)
		{
			InitializeComponent();
            MFRM = motherFrm;
            //InitializeNiceMenu(); //Modify our menues to ones with icons using NiceMenu class
            autoOpenDatabase();   //load database
		}

		#region Windows Form Designer generated code
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

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(IncomingCallsFrm));
            this.imageMenu = new System.Windows.Forms.ImageList(this.components);
            this.menuFile = new System.Windows.Forms.MainMenu(this.components);
            this.menuF_open = new System.Windows.Forms.MenuItem();
            this.dGrid = new System.Windows.Forms.DataGrid();
            this.comboTables = new System.Windows.Forms.ComboBox();
            this.lblNavLocation = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.Findbtn = new System.Windows.Forms.Button();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton3 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton5 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton6 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton8 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton9 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton7 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton4 = new System.Windows.Forms.ToolStripButton();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.richTextBox2 = new System.Windows.Forms.RichTextBox();
            this.richTextBox3 = new System.Windows.Forms.RichTextBox();
            this.richTextBox4 = new System.Windows.Forms.RichTextBox();
            this.richTextBox5 = new System.Windows.Forms.RichTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnNavPrev = new System.Windows.Forms.Button();
            this.btnNavNext = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mailMergeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dGrid)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // imageMenu
            // 
            this.imageMenu.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageMenu.ImageSize = new System.Drawing.Size(16, 16);
            this.imageMenu.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // menuF_open
            // 
            this.menuF_open.Index = -1;
            this.menuF_open.Text = "";
            // 
            // dGrid
            // 
            this.dGrid.BackgroundColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.dGrid.DataMember = "";
            this.dGrid.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dGrid.Location = new System.Drawing.Point(0, 337);
            this.dGrid.MinimumSize = new System.Drawing.Size(0, 200);
            this.dGrid.Name = "dGrid";
            this.dGrid.ReadOnly = true;
            this.dGrid.Size = new System.Drawing.Size(1122, 200);
            this.dGrid.TabIndex = 58;
            this.dGrid.CurrentCellChanged += new System.EventHandler(this.dGrid_CurrentCellChanged);
            this.dGrid.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dGrid_MouseDown);
            // 
            // comboTables
            // 
            this.comboTables.Location = new System.Drawing.Point(12, 292);
            this.comboTables.Name = "comboTables";
            this.comboTables.Size = new System.Drawing.Size(201, 24);
            this.comboTables.TabIndex = 68;
            this.comboTables.Text = "Tables";
            this.comboTables.Visible = false;
            // 
            // lblNavLocation
            // 
            this.lblNavLocation.BackColor = System.Drawing.Color.White;
            this.lblNavLocation.Location = new System.Drawing.Point(268, 292);
            this.lblNavLocation.MinimumSize = new System.Drawing.Size(0, 21);
            this.lblNavLocation.Name = "lblNavLocation";
            this.lblNavLocation.Size = new System.Drawing.Size(95, 21);
            this.lblNavLocation.TabIndex = 63;
            this.lblNavLocation.Text = "No Records";
            this.lblNavLocation.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblNavLocation.Visible = false;
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.Color.YellowGreen;
            this.btnAdd.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnAdd.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Green;
            this.btnAdd.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAdd.Location = new System.Drawing.Point(492, 292);
            this.btnAdd.MinimumSize = new System.Drawing.Size(0, 21);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 21);
            this.btnAdd.TabIndex = 65;
            this.btnAdd.Text = "&Add New";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Visible = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.btnDelete.Cursor = System.Windows.Forms.Cursors.No;
            this.btnDelete.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Red;
            this.btnDelete.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Fuchsia;
            this.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDelete.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDelete.Location = new System.Drawing.Point(573, 292);
            this.btnDelete.MinimumSize = new System.Drawing.Size(0, 21);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 21);
            this.btnDelete.TabIndex = 66;
            this.btnDelete.Text = "&Delete";
            this.btnDelete.UseVisualStyleBackColor = false;
            this.btnDelete.Visible = false;
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.Color.Salmon;
            this.btnCancel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCancel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Red;
            this.btnCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(654, 292);
            this.btnCancel.MinimumSize = new System.Drawing.Size(0, 21);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 21);
            this.btnCancel.TabIndex = 67;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.btnUpdate.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnUpdate.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnUpdate.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Lime;
            this.btnUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnUpdate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpdate.Location = new System.Drawing.Point(735, 292);
            this.btnUpdate.MinimumSize = new System.Drawing.Size(0, 21);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(75, 21);
            this.btnUpdate.TabIndex = 61;
            this.btnUpdate.Text = "Save";
            this.btnUpdate.UseVisualStyleBackColor = false;
            this.btnUpdate.Visible = false;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // Findbtn
            // 
            this.Findbtn.BackColor = System.Drawing.Color.OrangeRed;
            this.Findbtn.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.Findbtn.Location = new System.Drawing.Point(654, 316);
            this.Findbtn.Name = "Findbtn";
            this.Findbtn.Size = new System.Drawing.Size(75, 21);
            this.Findbtn.TabIndex = 70;
            this.Findbtn.Text = "&Find";
            this.Findbtn.UseVisualStyleBackColor = false;
            this.Findbtn.Visible = false;
            this.Findbtn.Click += new System.EventHandler(this.Findbtn_Click);
            // 
            // toolStrip1
            // 
            this.toolStrip1.BackColor = System.Drawing.SystemColors.ControlDark;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel1,
            this.toolStripButton1,
            this.toolStripButton2,
            this.toolStripButton3,
            this.toolStripButton5,
            this.toolStripButton6,
            this.toolStripButton8,
            this.toolStripButton9,
            this.toolStripButton7,
            this.toolStripButton4});
            this.toolStrip1.Location = new System.Drawing.Point(0, 28);
            this.toolStrip1.MinimumSize = new System.Drawing.Size(0, 26);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1122, 26);
            this.toolStrip1.Stretch = true;
            this.toolStrip1.TabIndex = 0;
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(126, 23);
            this.toolStripLabel1.Text = "Update Records:  ";
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.White;
            this.toolStripButton1.Margin = new System.Windows.Forms.Padding(0, 1, 1, 1);
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(23, 24);
            this.toolStripButton1.Text = "Previous Record";
            this.toolStripButton1.Click += new System.EventHandler(this.btnNavPrev_Click);
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton2.Image")));
            this.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton2.Margin = new System.Windows.Forms.Padding(1);
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(23, 24);
            this.toolStripButton2.Text = "Next Record";
            this.toolStripButton2.Click += new System.EventHandler(this.btnNavNext_Click);
            // 
            // toolStripButton3
            // 
            this.toolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton3.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton3.Image")));
            this.toolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton3.Margin = new System.Windows.Forms.Padding(0, 1, 2, 1);
            this.toolStripButton3.Name = "toolStripButton3";
            this.toolStripButton3.Size = new System.Drawing.Size(23, 24);
            this.toolStripButton3.Text = "Add New Record";
            this.toolStripButton3.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // toolStripButton5
            // 
            this.toolStripButton5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton5.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton5.Image")));
            this.toolStripButton5.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton5.Margin = new System.Windows.Forms.Padding(0, 1, 2, 1);
            this.toolStripButton5.Name = "toolStripButton5";
            this.toolStripButton5.Size = new System.Drawing.Size(23, 24);
            this.toolStripButton5.Text = "Cancel";
            this.toolStripButton5.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // toolStripButton6
            // 
            this.toolStripButton6.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton6.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton6.Image")));
            this.toolStripButton6.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton6.Margin = new System.Windows.Forms.Padding(0, 1, 2, 1);
            this.toolStripButton6.Name = "toolStripButton6";
            this.toolStripButton6.Size = new System.Drawing.Size(23, 24);
            this.toolStripButton6.Text = "Save";
            this.toolStripButton6.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // toolStripButton8
            // 
            this.toolStripButton8.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton8.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton8.Image")));
            this.toolStripButton8.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton8.Margin = new System.Windows.Forms.Padding(0, 1, 2, 1);
            this.toolStripButton8.Name = "toolStripButton8";
            this.toolStripButton8.Size = new System.Drawing.Size(23, 24);
            this.toolStripButton8.Text = "Change background color";
            this.toolStripButton8.Click += new System.EventHandler(this.toolStripButton8_Click);
            // 
            // toolStripButton9
            // 
            this.toolStripButton9.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton9.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton9.Image")));
            this.toolStripButton9.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton9.Name = "toolStripButton9";
            this.toolStripButton9.Size = new System.Drawing.Size(23, 23);
            this.toolStripButton9.Text = "Load current data from network. This action may overwrite your unsaved changes.";
            this.toolStripButton9.Click += new System.EventHandler(this.btn_Refresh);
            // 
            // toolStripButton7
            // 
            this.toolStripButton7.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton7.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStripButton7.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton7.Image")));
            this.toolStripButton7.ImageTransparentColor = System.Drawing.Color.White;
            this.toolStripButton7.Margin = new System.Windows.Forms.Padding(0, 1, 2, 1);
            this.toolStripButton7.Name = "toolStripButton7";
            this.toolStripButton7.Size = new System.Drawing.Size(23, 24);
            this.toolStripButton7.Text = "Help";
            this.toolStripButton7.Click += new System.EventHandler(this.toolStripButton7_Click);
            // 
            // toolStripButton4
            // 
            this.toolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton4.Enabled = false;
            this.toolStripButton4.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton4.Image")));
            this.toolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton4.Margin = new System.Windows.Forms.Padding(0, 1, 2, 1);
            this.toolStripButton4.Name = "toolStripButton4";
            this.toolStripButton4.Size = new System.Drawing.Size(23, 24);
            this.toolStripButton4.Text = "Delete";
            this.toolStripButton4.Visible = false;
            // 
            // textBox2
            // 
            this.textBox2.ForeColor = System.Drawing.Color.SkyBlue;
            this.textBox2.Location = new System.Drawing.Point(973, 317);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(75, 22);
            this.textBox2.TabIndex = 71;
            this.textBox2.Text = "search field";
            this.textBox2.MouseClick += new System.Windows.Forms.MouseEventHandler(this.TextControl_MouseClickTextBox);
            this.textBox2.ModifiedChanged += new System.EventHandler(this.TextControl_Modified);
            this.textBox2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.CheckKeys);
            // 
            // richTextBox2
            // 
            this.richTextBox2.Location = new System.Drawing.Point(109, 68);
            this.richTextBox2.Name = "richTextBox2";
            this.richTextBox2.Size = new System.Drawing.Size(254, 25);
            this.richTextBox2.TabIndex = 72;
            this.richTextBox2.Text = "";
            // 
            // richTextBox3
            // 
            this.richTextBox3.Location = new System.Drawing.Point(109, 108);
            this.richTextBox3.Name = "richTextBox3";
            this.richTextBox3.Size = new System.Drawing.Size(254, 25);
            this.richTextBox3.TabIndex = 73;
            this.richTextBox3.Text = "";
            // 
            // richTextBox4
            // 
            this.richTextBox4.Location = new System.Drawing.Point(109, 148);
            this.richTextBox4.Name = "richTextBox4";
            this.richTextBox4.Size = new System.Drawing.Size(162, 25);
            this.richTextBox4.TabIndex = 74;
            this.richTextBox4.Text = "";
            // 
            // richTextBox5
            // 
            this.richTextBox5.Location = new System.Drawing.Point(109, 190);
            this.richTextBox5.Name = "richTextBox5";
            this.richTextBox5.Size = new System.Drawing.Size(162, 25);
            this.richTextBox5.TabIndex = 75;
            this.richTextBox5.Text = "";
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Location = new System.Drawing.Point(99, 220);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(900, 2);
            this.label1.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 210);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 17);
            this.label2.TabIndex = 76;
            this.label2.Text = "Origination";
            // 
            // btnNavPrev
            // 
            this.btnNavPrev.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNavPrev.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.btnNavPrev.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.btnNavPrev.Image = ((System.Drawing.Image)(resources.GetObject("btnNavPrev.Image")));
            this.btnNavPrev.Location = new System.Drawing.Point(236, 292);
            this.btnNavPrev.MinimumSize = new System.Drawing.Size(0, 21);
            this.btnNavPrev.Name = "btnNavPrev";
            this.btnNavPrev.Size = new System.Drawing.Size(35, 21);
            this.btnNavPrev.TabIndex = 62;
            this.btnNavPrev.Visible = false;
            this.btnNavPrev.Click += new System.EventHandler(this.btnNavPrev_Click);
            // 
            // btnNavNext
            // 
            this.btnNavNext.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNavNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNavNext.Image")));
            this.btnNavNext.Location = new System.Drawing.Point(364, 292);
            this.btnNavNext.MinimumSize = new System.Drawing.Size(0, 21);
            this.btnNavNext.Name = "btnNavNext";
            this.btnNavNext.Size = new System.Drawing.Size(35, 21);
            this.btnNavNext.TabIndex = 64;
            this.btnNavNext.Visible = false;
            this.btnNavNext.Click += new System.EventHandler(this.btnNavNext_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.toolsToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1122, 28);
            this.menuStrip1.TabIndex = 77;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveToolStripMenuItem,
            this.closeToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(44, 24);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(114, 24);
            this.saveToolStripMenuItem.Text = "Save";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(114, 24);
            this.closeToolStripMenuItem.Text = "Close";
            this.closeToolStripMenuItem.Click += new System.EventHandler(this.closeToolStripMenuItem_Click);
            // 
            // toolsToolStripMenuItem
            // 
            this.toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mailMergeToolStripMenuItem});
            this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            this.toolsToolStripMenuItem.Size = new System.Drawing.Size(57, 24);
            this.toolsToolStripMenuItem.Text = "Tools";
            // 
            // mailMergeToolStripMenuItem
            // 
            this.mailMergeToolStripMenuItem.Name = "mailMergeToolStripMenuItem";
            this.mailMergeToolStripMenuItem.Size = new System.Drawing.Size(154, 24);
            this.mailMergeToolStripMenuItem.Text = "Mail Merge";
            this.mailMergeToolStripMenuItem.Click += new System.EventHandler(this.mailMergeToolStripMenuItem_Click);
            // 
            // IncomingCallsFrm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoScroll = true;
            this.AutoValidate = System.Windows.Forms.AutoValidate.Disable;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.ClientSize = new System.Drawing.Size(1143, 362);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.richTextBox5);
            this.Controls.Add(this.richTextBox4);
            this.Controls.Add(this.richTextBox3);
            this.Controls.Add(this.richTextBox2);
            this.Controls.Add(this.Findbtn);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.dGrid);
            this.Controls.Add(this.comboTables);
            this.Controls.Add(this.btnNavPrev);
            this.Controls.Add(this.lblNavLocation);
            this.Controls.Add(this.btnNavNext);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnUpdate);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Menu = this.menuFile;
            this.MinimumSize = new System.Drawing.Size(1098, 200);
            this.Name = "IncomingCallsFrm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Easy Master Client - Incoming Calls";
            this.Resize += new System.EventHandler(this.MainFrm_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dGrid)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

/*
		#region "Initialize Nice Menu"
		private void InitializeNiceMenu()
		{
			// Here the menues are defined as Nice Menu
			myNiceMenu = new NiceMenu();
			
			// The icons are attached Here
			myNiceMenu.MenuImages = imageMenu;

			//The click event is declared
			//myNiceMenu.UpdateMenu(this.menuFile, new NiceMenuClickEvent(menuClickEvent));
		}
		#endregion
*/

/*
		#region "Nice Menu Click Event"
		// Nice Menu Click Event
		// this event is fired when someone clicks an
		// item in the menu
		public void menuClickEvent(object sender, System.EventArgs e)
		{
			NiceMenu itemIncCalls = (NiceMenu)sender;

            switch (itemIncCalls.Text)
			{
				case "Save":
					//Save the changes to the source database the mdb file
					btnUpdate.Focus();
					DialogResult R = MessageBox.Show("Are you sure you want to save? Changes will be permenant.","Save Confirmation",MessageBoxButtons.YesNo,MessageBoxIcon.Exclamation);
                    if (R == DialogResult.Yes) btnUpdate_Click(sender, e);
					break;
                case "Close":
                    //Exit the application
                    //Check and Alert the user if data changed and is not saved!!
                    if (Check_If_Data_Changed() == true)
                    {
                        DialogResult r = MessageBox.Show("The database file changed, are you sure you want exit without saving?", "Exit Without Save", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                        if (r == DialogResult.Yes) //this.Close();     
                        {
                            this.Hide();
                            main(sender, e);
                        }
                    }
                    else //this.Close(); 
                    {
                        this.Hide();
                        StartFormOpen(sender, e);
                    }
                    break;
                case "About":
                    menuAboutClick();
                    break;
                case "Get Help":
                    //menuItem3_Click();
                    toolStripButton7_Click(sender, e);
                    break;
                case "Mail Merge":
                    if (DataLoaded == false)
                    {
                        MessageBox.Show("The database file has to be loaded.");
                        break;
                    }
                    if (Check_If_Data_Changed() == true)
                    {
                        DialogResult r = MessageBox.Show("The database file changed, are you sure you want to proceed?", "Go on without saving", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                        if (r == DialogResult.Yes) menuItem6_BuildDocMailMergeClick(sender, e);
                    }
                    else
                    {
                        menuItem6_BuildDocMailMergeClick(sender, e);
                    }
                    break;  
			}
		}
		#endregion
*/
        #region "Save"
        //This routine handles saving
		//data changes to the main file
		private void Save_File()
		{		
			// Create a new dataset to hold the changes that have
			// been made to the main dataset.
			DataSet objDataSetChanges = new DataSet();
			// Stop any current edits.
			this.BindingContext[accessDataSet,comboTables.Text].EndCurrentEdit();
			// Get the changes that have been made to the main dataset.
			objDataSetChanges = ((DataSet)(accessDataSet.GetChanges()));
			// Check to see if any changes have been made.
			if ((objDataSetChanges != null)) 
			{
				// There are changes that need to be made, so attempt to update the datasource by
				// calling the update method and passing the dataset and any parameters.
				UpdateDataSource(objDataSetChanges);

				//Make sure the database connection is closed!!
				if (this.accessConnection.State.ToString()!="Closed") this.accessConnection.Close();
			}
		}
		#endregion

		#region "Menu Open"
		// Menu Open Click routine!!
		public void menuOpenClick()
		{
            try
            {
            
			OpenFileDialog openFile = new OpenFileDialog();
			openFile.FileName = "";
			//// Make sure only *.mdb files can be opened
			//// by using a filter
            openFile.Filter = "Microsoft Access Application (*.mdb)| *.mdb";    //mdb

			System.Windows.Forms.DialogResult res = openFile.ShowDialog();
			if (res==System.Windows.Forms.DialogResult.Cancel) return;
           
			////Change the mouse icon and caption 
			////of the form to inform the user that
            ////the data is being loaded
			this.Cursor = Cursors.WaitCursor;
			this.Text = "MasterClient: Loading Data Please Wait...";
            
			////The connection parameters
			//this.accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " + openFile.FileName;
            ////accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=C:\\Law\\TEST-2_MasterClient\\Database61.mdb";
			string stDataSource = openFile.FileName;
			dataSourceFile = openFile.FileName;
			////remove any dunamically created controls from the form
			removeMadeControls();
           
            /*
            ///////////////////////
            string fbPath = Application.StartupPath;
            string fname = "Database61.mdb";
            //string filename = fbPath + @"\" + fname;
            string filename = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " + fbPath + @"\" + fname;
            //System.Diagnostics.Process.Start(filename);
            ///////////////////////
            */

			
				// Attempt to fill the temporary dataset.
				// Turn off constraint checking before the dataset is ed.
				// This allows the adapters to fill the dataset without concern
				// for dependencies between the tables.
				accessDataSet = new DataSet();
				accessDataSet.EnforceConstraints = false;

				((System.ComponentModel.ISupportInitialize)(this.accessDataSet)).BeginInit();
				
				try 
				{
                    //MessageBox.Show("Message 5");
					// Open the connection.
                    try
                    {
                        this.accessConnection.Open();
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show("Failed to connect to data source", "Warning");
                        MessageBox.Show("Connecting to data source from this system... Please click OK", "Warning");
                    }
					
                    //MessageBox.Show("Message 6");
					//Get how many tables this datafile has
                    DataTable schemaTable = this.accessConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
						new object[] {null, null, null, "TABLE"});

					//Update the comboBox to have all
					//the tables in the database
					//and keep the first table in the database
					//as the combobox's text
					comboTables.Items.Clear();

					int i = 0;

					foreach (DataRow r in schemaTable.Rows)
					{
						if (i==0) 
						{
							comboTables.Text=(r["TABLE_NAME"].ToString());
							comboTables.Items.Add(r["TABLE_NAME"].ToString());
						}
						else comboTables.Items.Add(r["TABLE_NAME"].ToString());
						i++;
					}
                    //MessageBox.Show("Message 7");
					//load data
                    comboTables.Text = "Incoming Calls";
					ComboBoxText = comboTables.Text;
					//MessageBox.Show("Message 8");
					//Call the LoadData Routine!!
					LoadData("Select * From [" + comboTables.Text + "]");
                    //LoadData("Select * From [Incoming Calls]");
                    
				}
				catch (System.Data.OleDb.OleDbException fillException) 
				{
					//report error incase of failure in loading data
					MessageBox.Show(fillException.Message);
				}
				finally 
				{
					// Turn constraint checking back on.
					accessDataSet.EnforceConstraints = true;
					// Close the connection whether or not the exception was thrown.
                    this.accessConnection.Close();
				}
			}
			catch (System.Data.OleDb.OleDbException eFillDataSet) 
			{
                MessageBox.Show("Exception loading the database");
				throw eFillDataSet;
			}

			//Return the cursor and form's caption to their normal state
			this.Cursor = Cursors.Arrow;
            this.Text = "Easy Master Client - Incoming Calls";
		}
		#endregion

        #region "Auto Open Database"
        // load database automaticaly at start up
        public void autoOpenDatabase()
        {
            try
            {

                OpenFileDialog openFile = new OpenFileDialog();
                openFile.FileName = "";
                //// Make sure only *.mdb files can be opened
                //// by using a filter
                openFile.Filter = "Microsoft Access Application (*.mdb)| *.mdb";    //mdb

                //System.Windows.Forms.DialogResult res = openFile.ShowDialog();
                //if (res == System.Windows.Forms.DialogResult.Cancel) return;

                ////Change the mouse icon and caption 
                ////of the form to inform the user that
                ////the data is being loaded
                this.Cursor = Cursors.WaitCursor;
                this.Text = "MasterClient: Loading Data Please Wait...";

                ////The connection parameters
                //accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=C:\\Misc\\Test\\EasyClientMaster_02092013\\Database63_be.mdb";
                //this.accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=Database63_fe.mdb";
                string executable = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string path = (System.IO.Path.GetDirectoryName(executable));
                AppDomain.CurrentDomain.SetData("DataDirectory", path);
                this.accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=|DataDirectory|\Database63_fe.mdb";

                //accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=C:\\Users\\Mikhail\\Documents\\NEW MCL\\Database63_be.mdb";

                
                store_accessConnection = this.accessConnection;
                //accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=Database63_be.mdb";
                string stDataSource = openFile.FileName;
                dataSourceFile = openFile.FileName;
                ////remove any dunamically created controls from the form
                removeMadeControls();


                // Attempt to fill the temporary dataset.
                // Turn off constraint checking before the dataset is ed.
                // This allows the adapters to fill the dataset without concern
                // for dependencies between the tables.
                accessDataSet = new DataSet();
                accessDataSet.EnforceConstraints = false;

                ((System.ComponentModel.ISupportInitialize)(this.accessDataSet)).BeginInit();

                try
                {
                    //MessageBox.Show("Message 5");
                    // Open the connection.
                    //accessConnection.Open();
                    try
                    {
                        accessConnection.Open();
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show("Failed to connect to data source", "Warning");
                        MessageBox.Show("Connecting to data source from this system... Please click OK", "Warning");
                    }

                    //MessageBox.Show("Message 6");
                    //Get how many tables this datafile has
                    DataTable schemaTable = this.accessConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                        new object[] { null, null, null, "TABLE" });

                    //Update the comboBox to have all
                    //the tables in the database
                    //and keep the first table in the database
                    //as the combobox's text
                    comboTables.Items.Clear();

                    int i = 0;

                    foreach (DataRow r in schemaTable.Rows)
                    {
                        if (i == 0)
                        {
                            comboTables.Text = (r["TABLE_NAME"].ToString());
                            comboTables.Items.Add(r["TABLE_NAME"].ToString());
                        }
                        else comboTables.Items.Add(r["TABLE_NAME"].ToString());
                        i++;
                    }
                    //MessageBox.Show("Message 7");
                    //load data
                    comboTables.Text = "Incoming Calls";
                    ComboBoxText = comboTables.Text;
                    //MessageBox.Show("Message 8");

                    //Call the LoadData Routine!!
                    LoadData("Select * From [" + comboTables.Text + "]");
                    //LoadData("Select * From [Incoming Calls]");
                }
                catch (System.Data.OleDb.OleDbException fillException)
                {
                    //report error incase of failure in loading data
                    MessageBox.Show(fillException.Message);
                }
                finally
                {
                    // Turn constraint checking back on.
                    accessDataSet.EnforceConstraints = true;
                    // Close the connection whether or not the exception was thrown.
                    this.accessConnection.Close();
                }
            }
            catch (System.Data.OleDb.OleDbException eFillDataSet)
            {
                MessageBox.Show("Exception loading the database");
                throw eFillDataSet;
            }
            //Return the cursor and form's caption to their normal state
            this.Cursor = Cursors.Arrow;
            this.Text = "Easy Master Client - Incoming Calls";
        }
        #endregion

		#region "Menu Search Click"
		//Search Menu is clicked
		private void menuSearchClick()
		{
			//Create a new instance of the search form
			//Specifying the datasource and table to view
			Form sfrm =  new searchfrm(this,dataSourceFile, "[" +comboTables.Text +"]" ,"Select * From [" + comboTables.Text +" ]");
			// Create a new dataset to hold the changes that have been made to the main dataset.
			DataSet objDataSetChanges = new DataSet();
			// Stop any current edits.
			this.BindingContext[accessDataSet,comboTables.Text].EndCurrentEdit();
			// Get the changes that have been made to the main dataset.
			objDataSetChanges = ((DataSet)(accessDataSet.GetChanges()));
			// Check to see if any changes have been made.
			if ((objDataSetChanges != null)) 
			{
				//alert the user that in order to view the same data
				//in the search he/she has to save the file
				DialogResult r  = MessageBox.Show("The database file changed, in oder to see your changes in the search form you have to save, continue any way?","Change In Data File Detected",MessageBoxButtons.YesNo,MessageBoxIcon.Exclamation);
				if (r==DialogResult.No) return;
			}
			Form.ActiveForm.Enabled = false;
			sfrm.Show();
		}
		#endregion

        #region "Menu Incoming Calls Click"
        //Incoming Calls Menu is clicked
        private void menuIncomingCallsClick(object sender, EventArgs e)
        {
            //Create a new instance of the search form
            //Specifying the datasource and table to view
            Form sfrm = new searchfrm(this, dataSourceFile, "[" + comboTables.Text + "]", "Select * From [" + comboTables.Text + " ]");
            // Create a new dataset to hold the changes that have been made to the main dataset.
            DataSet objDataSetChanges = new DataSet();
            // Stop any current edits.
            this.BindingContext[accessDataSet, comboTables.Text].EndCurrentEdit();
            // Get the changes that have been made to the main dataset.
            objDataSetChanges = ((DataSet)(accessDataSet.GetChanges()));
            // Check to see if any changes have been made.
            if ((objDataSetChanges != null))
            {
                //alert the user that in order to view the same data
                //in the search he/she has to save the file
                DialogResult r = MessageBox.Show("The database file changed, in oder to see your changes in the search form you have to save, continue any way?", "Change In Data File Detected", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (r == DialogResult.No) return;
            }
            Form.ActiveForm.Enabled = false;
            sfrm.Show();
        }
        #endregion

        #region "Menu About Click"
        private void menuAboutClick()
        {
            Form abtBox = new About();           
            abtBox.Show();
        }
        #endregion
 
        #region "LoadData"
        //Here data is read from the database
		//and all the parameters are set
		//Select string is the SQL command required 
		//to view the data
		public void LoadData(string SelectString)
		{
            //dGrid.AllowSorting = false;

            tableName = "[" + comboTables.Text + "]";  //DataBase table name
			this.Cursor = Cursors.WaitCursor;
            this.Text = "MasterClient: Loading Data Please Wait...";
			try	
			{
                if ( !accessDataSet.IsInitialized ) accessDataSet = new DataSet();
                //accessDataSet = new DataSet();

				DataTable accessDataTable;
				//create new instances for select, insert and update
				//and delete commands to be used with the adapter
				OleDbCommand accessSelectCommand = new OleDbCommand();
				OleDbCommand accessInsertCommand = new OleDbCommand();
				OleDbCommand accessUpdateCommand = new OleDbCommand();
				OleDbCommand accessDeleteCommand = new OleDbCommand();
				
				accessDataAdapter = null;
				accessDataAdapter = new OleDbDataAdapter();

				accessSelectCommand.CommandText = SelectString;
				accessSelectCommand.Connection =  accessConnection;
				accessDataAdapter.SelectCommand = accessSelectCommand;


                

				// Attempt to fill the dataset through the accessDataAdapter
				accessDataAdapter.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {																								   
																									   new System.Data.Common.DataTableMapping("Table", comboTables.Text)});

                
				//populate the DataSet with existing constraints information from a data source
                accessDataAdapter.FillSchema(accessDataSet, SchemaType.Source, comboTables.Text);


            
                //accessDataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;   ///////////////added///////////


				// Fill the dataset
                accessDataAdapter.Fill(accessDataSet);
                //accessDataAdapter.Fill(accessDataSet, comboTables.Text);
                


				//create an instance for a datatable
                //accessDataSet.Tables[comboTables.Text].PrimaryKey = new DataColumn[] { accessDataSet.Tables[comboTables.Text].Columns[0] }; ///////////////added//////////////
                accessDataTable = accessDataSet.Tables[comboTables.Text];

     			
				// Make dynamic Insert Commands and Parameters
				accessInsertCommand.Connection = accessConnection;
				make_Insert_Command(accessDataTable,accessInsertCommand);
				accessDataAdapter.InsertCommand = accessInsertCommand;
 
				// The dynamic Update Commands and Parameters
				accessUpdateCommand.Connection = accessConnection;			
				make_Update_Command(accessDataTable,accessUpdateCommand);
				accessDataAdapter.UpdateCommand = accessUpdateCommand;

				// The dynamic Delete Commands and Parameters
				accessDeleteCommand.Connection = accessConnection;			
				make_Delete_Command(accessDataTable,accessDeleteCommand);
				accessDataAdapter.DeleteCommand = accessDeleteCommand;

				// Dynamic Controls Postions
				int controlTop = 10;
				int controlLeft = 10;
                string readColumnHeader = "";
				
				//Get all the System.DataTypes of all the
				//columns in the table and assign them to the
				//array colType
				colType = new string[accessDataTable.Columns.Count];

				//Here AutoMenu is created which would allow the user
				//to insert automatic incrementation of numbers (+1 on the
				//last cell) or insert today's date for datetime type
				//columns
				AutoMenu = new ContextMenu[accessDataTable.Columns.Count];

				//Create dynamically all the textboxes and labels
				//which will hold and link information to the database
				//making it easier to input data
                
				for (int i=0;i< accessDataTable.Columns.Count;i++)
				{
                    //Find all columns and put them in the combobox

                    if (i == 0)  
                    {
                        controlTop = controlTop + 20;
                    }
					colType[i] = accessDataTable.Columns[i].DataType.ToString();

					//Create the control (Label)

					Label LabelControl = (Label)CreateControls.MakeControl("Label",30,100,
						controlLeft,controlTop+3,
						accessDataTable.Columns[i].Caption + " :","cLabel"+i);

                    readColumnHeader = accessDataTable.Columns[i].Caption;

                    //Create comboBoxes (dropdown list menus) for specific columns.
                    //Otherwise creare normal Rich TextBoxes for data display and edit.
                    if (readColumnHeader == "ID")
                    {

                        ////Bind the textbox control to the database table column
                        //richTextBox1.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));

                        ////Finally add the controls to the form
                        //this.Controls.Add(LabelControl);

                        controlTop += LabelControl.Height + 10;
                    }
                    else if (readColumnHeader == "First Name")
                    {
                        //Create the control (TextBox)

                        richTextBox2.Leave += new EventHandler(HasSpecialChars);

                        richTextBox2.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));
                        richTextBox2.DataBindings["Text"].FormatString = "First Name";
                        richTextBox2.DataBindings["Text"].NullValue = "";
                        richTextBox2.DataBindings["Text"].FormattingEnabled = true;
                        richTextBox2.DataBindings["Text"].BindingComplete +=
                                    delegate(object sender, BindingCompleteEventArgs e)
                                    {
                                        if (e.Exception is FormatException)
                                            MessageBox.Show("This entry can not be saved. Wrong formating, should be: " + richTextBox2.DataBindings["Text"].FormatString,
                                                                "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        if (e.BindingCompleteState != BindingCompleteState.Success)
                                            MessageBox.Show("partNumberBinding: " + e.ErrorText,
                                                                "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    };

                        richTextBox2.DoubleClick += new EventHandler(View_Same_Test_in_Bigger_Window);
                        //this.label2.Text = LabelControl.Text;

                        //Finally add the controls to the form
                        //this.Controls.Add(richTextBox2);
                        this.Controls.Add(LabelControl);

                        controlTop += LabelControl.Height + 10;
                    }
                    else if (readColumnHeader == "Last Name")
                    {
                        //Create the control (TextBox)

                        richTextBox3.Leave += new EventHandler(HasSpecialChars);

                        richTextBox3.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));
                        richTextBox3.DataBindings["Text"].FormatString = "Last Name";
                        richTextBox3.DataBindings["Text"].NullValue = "";
                        richTextBox3.DataBindings["Text"].FormattingEnabled = true;
                        richTextBox3.DataBindings["Text"].BindingComplete +=
                                    delegate(object sender, BindingCompleteEventArgs e)
                                    {
                                        if (e.Exception is FormatException)
                                            MessageBox.Show("This entry can not be saved. Wrong formating, should be: " + richTextBox3.DataBindings["Text"].FormatString,
                                                                "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        if (e.BindingCompleteState != BindingCompleteState.Success)
                                            MessageBox.Show("partNumberBinding: " + e.ErrorText,
                                                                "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    };

                        richTextBox3.DoubleClick += new EventHandler(View_Same_Test_in_Bigger_Window);

                        //Finally add the controls to the form
                        this.Controls.Add(LabelControl);

                        controlTop += LabelControl.Height + 10;
                    }
                    else if (readColumnHeader == "Phone Number")
                    {
                        //Create the control (TextBox)

                        richTextBox4.Leave += new EventHandler(HasSpecialChars);

                        richTextBox4.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));
                        richTextBox4.DataBindings["Text"].FormatString = "Phone Number";
                        richTextBox4.DataBindings["Text"].NullValue = "";
                        richTextBox4.DataBindings["Text"].FormattingEnabled = true;
                        richTextBox4.DataBindings["Text"].BindingComplete +=
                                    delegate(object sender, BindingCompleteEventArgs e)
                                    {
                                        if (e.BindingCompleteState != BindingCompleteState.Success)
                                            MessageBox.Show("partNumberBinding: " + e.ErrorText,
                                                                "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    };

                        //Finally add the controls to the form
                        this.Controls.Add(LabelControl);

                        controlTop += LabelControl.Height + 10;
                    }
                    else if (readColumnHeader == "E-mail")
                    {
                        //Create the control (TextBox)

                        richTextBox5.Leave += new EventHandler(HasSpecialChars);

                        richTextBox5.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));
                        richTextBox5.DataBindings["Text"].FormatString = "E-mail";
                        richTextBox5.DataBindings["Text"].NullValue = "";
                        richTextBox5.DataBindings["Text"].FormattingEnabled = true;
                        richTextBox5.DataBindings["Text"].BindingComplete +=
                                    delegate(object sender, BindingCompleteEventArgs e)
                                    {
                                        if (e.BindingCompleteState != BindingCompleteState.Success)
                                            MessageBox.Show("partNumberBinding: " + e.ErrorText,
                                                                "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    };

                        richTextBox5.DoubleClick += new EventHandler(Open_Outlook_email);

                        //Finally add the controls to the form
                        this.Controls.Add(LabelControl);

                        controlTop += LabelControl.Height + 26;
                    }
                    else if (readColumnHeader == "Address")
                    {
                        // Create label
                        Label LabelControl_Address = (Label)CreateControls.MakeControl("Label", 30, 100,            //size (y, x)
                            controlLeft + 420, 70,                                                                  //position (from left, from top)
                            readColumnHeader + " :", "cLabel" + i);

                        //Create the control (TextBox)
                        RichTextBox TextControl = (RichTextBox)CreateControls.MakeControl("RichTextBox", 110, 300,  //size (y, x)
                            controlLeft + 500, 65,                                                                  //position (from left, from top)
                            readColumnHeader, "cText" + i);

                        //Assign the TabIndex sequentially to the created
                        //textbox control
                        TextControl.TabIndex = i;

                        TextControl.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));
                        TextControl.DoubleClick += new EventHandler(View_Same_Test_in_Bigger_Window);

                        //Finally add the controls to the form
                        this.Controls.Add(TextControl);
                        this.Controls.Add(LabelControl_Address);

                        //controlTop = controlTop - 30;    //readjust the position on the Form for the next field
                    }
                    else if (readColumnHeader == "Lawyer")
                    {
                        //Create the control (TextBox)
                        ComboBox TextControl = (ComboBox)CreateControls.MakeControl("ComboBox", 25, 162,
                            controlLeft + 100, controlTop,
                            readColumnHeader, "cText" + i);

                        //Assign the TabIndex sequentially to the created textbox control
                        TextControl.TabIndex = i;

                        //*  Create a dropdown list (comboBox) of unique dataset from the 
                        //*  specific column (Title: "Lawyer"). In other words,
                        //*  display in the list only unique values - do not repeate identical once.
                        DataTable dtUniqueCities = accessDataTable.DefaultView.ToTable(true, readColumnHeader);

                        dtUniqueCities.DefaultView.Sort = readColumnHeader;

                        TextControl.DataSource = dtUniqueCities;
                        TextControl.DisplayMember = readColumnHeader;
                        TextControl.ValueMember = "ID";
                        TextControl.MaxDropDownItems = 10;

                        // Bind the SelectedValueChanged event to our handler for it.
                        // This method is called when selected an item from comboBox list. Add that value to grid table.                        
                        TextControl.SelectionChangeCommitted += new EventHandler(ComboBox1_SelectedValueChanged);

                        //Bind the textbox control to the database table column
                        TextControl.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));

                        //Finally add the controls to the form
                        this.Controls.Add(TextControl);
                        this.Controls.Add(LabelControl);
                    }
                    else if (readColumnHeader == "Advice")
                    {
                        // Create label
                        Label LabelControl_Notes = (Label)CreateControls.MakeControl("Label", 30, 100,    //size (y, x)
                            controlLeft + 420, 251,                                                       //position (from left, from top)
                            readColumnHeader + " :", "cLabel" + i);

                        //Create the control (TextBox)
                        RichTextBox TextControl = (RichTextBox)CreateControls.MakeControl("RichTextBox", 150, 400,     //size (y, x)
                            controlLeft + 500, 245,                                                                    //position from left, from top)
                            readColumnHeader, "cText" + i);

                        //Assign the TabIndex sequentially to the created
                        //textbox control
                        TextControl.TabIndex = i;

                        TextControl.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));
                        TextControl.DoubleClick += new EventHandler(View_Same_Test_in_Bigger_Window);

                        //Finally add the controls to the form
                        this.Controls.Add(TextControl);
                        this.Controls.Add(LabelControl_Notes);

                        //controlTop = controlTop + 26;    //readjust the position on the Form for the next field
                    }
                    else
                    {
                        //Create the control (TextBox)
                        RichTextBox TextControl = (RichTextBox)CreateControls.MakeControl("RichTextBox", 25, 162,
                            controlLeft + 100, controlTop,
                            //controlLeft, controlTop+10,
                            readColumnHeader, "cText" + i);

                        //Assign the TabIndex sequentially to the created
                        //textbox control
                        TextControl.TabIndex = i;
////////////////////
                        if (colType[i] == "System.DateTime")
                        {                   
                            //UpdateTextBoxes(i, System.DateTime.Now.Date.ToShortDateString());                         
                            //TextControl.Text = System.DateTime.Now.Date.ToShortDateString();
                            //TextControl.Text = DateTime.Parse(TextControl.Text).ToString("yyyy-MM-dd");

                            TextControl.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));
                            //string temp1 = DateTime.Parse(TextControl.Text).ToString("yyyy-MM-dd");
                            TextControl.DataBindings["Text"].FormatString = "dd/MM/yyyy";
                            TextControl.DataBindings["Text"].NullValue = "";
                            TextControl.DataBindings["Text"].FormattingEnabled = true;
                            TextControl.DataBindings["Text"].BindingComplete +=
                                    delegate(object sender, BindingCompleteEventArgs e)
                                    {
                                        if (e.Exception is FormatException)
                                            MessageBox.Show("This entry can not be saved. Wrong formating, should be: " + TextControl.DataBindings["Text"].FormatString , 
                                                                "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    };

                        }
                        else
                        {
                            //Bind the textbox control to the database table column
                            //TextControl.DoubleClick += new EventHandler(View_Same_Test_in_Bigger_Window);
                            TextControl.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));

                        }
////////////////////

                        TextControl.DoubleClick += new EventHandler(View_Same_Test_in_Bigger_Window);
                        //TextControl.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.accessDataSet, comboTables.Text + "." + accessDataSet.Tables[comboTables.Text].Columns[i]));
                        
                        //Finally add the controls to the form
                        this.Controls.Add(TextControl);
                        this.Controls.Add(LabelControl);

                        controlTop += LabelControl.Height + 10;
                    }
 					
                    //check for exeptions
                    //this.textBox1 = TextControl;
                    //this.textBox1.TextChanged += new System.EventHandler(TextControl_TextChanged);
             
					//Here we arrange the controls on the form
					if (i == accessDataTable.Columns.Count-1) break;
/*					
					if (controlLeft == 10)
                    {
                        controlLeft = 220;
                    }
					else if ((controlLeft==220) &&  (accessDataTable.Columns.Count>6))
					{
						controlLeft=430;
						////this.Width = 1080;
					}
                    else if ((controlLeft == 430) && (accessDataTable.Columns.Count > 6))
                    {
                        controlLeft = 640;
                    }
                    else if ((controlLeft == 640) && (accessDataTable.Columns.Count > 6))
                    {
                        controlLeft = 850;
                    }
                    else
					{
						controlTop += LabelControl.Height + 10;
						controlLeft = 10;
					}
*/
                    
				}

                //MessageBox.Show("controlTop, accessDataTable.Columns.Count: " + controlTop + ", " + accessDataTable.Columns.Count);
				arrange_Controls(controlTop, accessDataTable.Columns.Count);
               
                
              
				//Create AutoInsertion of Date for only 'DateTime' Type
				//columns and
				//AutoNumber to only 'Integer' type columns
                //(dGrid[i,clmn].ToString().IndexOf(searchValue, StringComparison.OrdinalIgnoreCase) != -1)
                
				for (int i = 0; i< accessDataTable.Columns.Count ;i++) 
				{
					if (colType[i] == "System.Int32")
					{
						MenuItem Item1 = new MenuItem("Automatic Incrementation");
						Item1.Click += new System.EventHandler(this.cMenuClick);
						AutoMenu[i] = new ContextMenu(new MenuItem[]{Item1});
					}
					else if (colType[i] =="System.DateTime")
					{
						MenuItem Item1 = new MenuItem("Automatic Insertion of Today's Date");
						Item1.Click += new System.EventHandler(this.cMenuClick);
						AutoMenu[i] = new ContextMenu(new MenuItem[]{Item1});
					}
				}






                accessDataTable.DefaultView.Sort = "ID desc";
                //accessDataTable.DefaultView.Sort = "ID asc";
                //accessDataTable.DefaultView.RowFilter = "ID=5";
                dGrid.DataSource = accessDataTable.DefaultView;

                //this.Controls.Add(dGrid);

                //refresh the database view in the grid
                dGrid.SetDataBinding(accessDataTable.DataSet, comboTables.Text);
                this.DataSet_PositionChanged();

                
			}
			// catch any errors and display them
			catch(System.Data.OleDb.OleDbException e)
			{MessageBox.Show(e.Message);}

			this.Cursor = Cursors.Arrow;
            this.Text = "Incoming Calls";

            this.Width = 1080;
            this.Height = 720;

            //dGrid.AllowSorting = false;
            //////////////////////////////////////////////////////////////////
            //DataGridView dataGridView1 = new DataGridView();
            //dGrid.DataSource = accessDataTable.DefaultView;
            // If oldColumn is null, then the DataGridView is not currently sorted. 
            //dataGridView1.Sort(dataGridView1.Columns["ID"], ListSortDirection.Ascending);
            /////////////////////////////////////////////////////////////////////////////////////
		}
		#endregion


        private void ComboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            ComboBox T = (ComboBox)sender;
            string getItemFromList;

            if (T.SelectedIndex != -1)
            {
                T.Text = T.SelectedValue.ToString();
            }
        }
  


        private static void PrintDataView(DataView dv)
        {
            string text = "";
            // Printing first DataRowView to demo that the row in the first index of the DataView changes depending on sort and filters
            //Console.WriteLine("First DataRowView value is '{0}'", dv[0]["ID"]);
            text = "First DataRowView value is '{0}'" + "\n";

            // Printing all DataRowViews 
            foreach (DataRowView drv in dv)
            {
                //Console.WriteLine("\t {0}", drv["ID"]);
                // Example #2: Write one string to a text file. 
                text = text + "\t {0}" + drv["ID"] + "\n";
                // WriteAllText creates a file, writes the specified string to the file, 
                // and then closes the file.
                

            }
            System.IO.File.WriteAllText(@"C:\Misc\Test\EasyClientMaster_20140420_2\WriteText.txt", text);
        }



        #region "load Data from Find results"
        //This routine loads results from 
        //Find into the DGrid
        private void loadData_Find(string SelectString)
        {
            accessDataSet.RejectChanges();
            accessDataSet.Clear();

            tableName = "[" + comboTables.Text + "]";  //DataBase table name

            OleDbCommand accessSelectCommand = new OleDbCommand();
            OleDbCommand accessInsertCommand = new OleDbCommand();
            OleDbDataAdapter accessDataAdapter = new OleDbDataAdapter();

            accessSelectCommand.CommandText = SelectString;
            accessSelectCommand.Connection = accessConnection;
            accessDataAdapter.SelectCommand = accessSelectCommand;

            // Attempt to fill the dataset through the OleDbDataAdapter1.
            accessDataAdapter.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
																								   new System.Data.Common.DataTableMapping("Table", tableName)});
            accessDataAdapter.Fill(accessDataSet);

            dGrid.SetDataBinding(accessDataSet, tableName);

            int col = (accessDataSet.Tables[tableName].Columns.Count);
            int row = (accessDataSet.Tables[tableName].Rows.Count);

            //if (doUpdate == true) checkedMenu = new String[col];
            elements = new object[col][];
            //FilterMenu = new ContextMenu[col];

            for (int i = 0; i < col; i++)
            {
                elements[i] = new object[row];
                //if (doUpdate == true) checkedMenu[i] = "None";
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

                    /*/http://www.codeproject.com/Questions/75660/How-to-sort-rows-if-datagrid
                    int cell = 0;
                    int num = dGrid[j, cell];

                    if (minValue > num)
                        {
                            minValue = num;
                            minValueRow = i;
                        }

                    if (maxValue < num)
                        {
                            maxValue = num;
                            maxValueRow = i;
                        }
                    //end of http://www.codeproject.com/Questions/75660/How-to-sort-rows-if-datagrid
                    */
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
                //lterMenu[i] = new ContextMenu();
                //make_menues(elements[i], FilterMenu[i]);
            }
            int rows = accessDataSet.Tables[tableName].Rows.Count;
            //textBox1.Text = rows.ToString();
        }
        #endregion

		#region "make_Insert_Command"
		//Here the command string and the parameters are
		//assigned dynamically for the insert command
		private void make_Insert_Command(DataTable datatable, OleDbCommand insertcommand)
		{
			string insertString1 = "INSERT INTO [" + comboTables.Text + "] (";
			string insertString2="";
			
			for (int i=0;i< datatable.Columns.Count;i++)
			{
				insertString1 += "[" + datatable.Columns[i].Caption + "]";  
				insertString2 += "?"; 
						
				if (i!= datatable.Columns.Count-1)
				{
					insertString1 +=",";
					insertString2 +=",";
				}
				else {insertString1 +=") VALUES (" + insertString2 + ")";}

				insertcommand.Parameters.Add(new System.Data.OleDb.OleDbParameter(datatable.Columns[i].Caption, System.Data.OleDb.OleDbType.Variant, 16, datatable.Columns[i].Caption));
			}
			insertcommand.CommandText = insertString1;
		}
		#endregion

		#region "make_Update_Command"
		//Here the command string and the parameters are
		//assigned for the update command
		private void make_Update_Command(DataTable datatable, OleDbCommand updatecommand)
		{
			string updateString1 = "UPDATE [" + comboTables.Text + "] SET ";
			string updateString2="";

			for (int i=0;i< datatable.Columns.Count;i++)
			{
				if (datatable.Columns[i].AutoIncrement!=true)
				{
					updateString1 += "[" + datatable.Columns[i].Caption + "]  = ? ";
					if (i!= datatable.Columns.Count-1) updateString1 +=" , ";
					OleDbType colSQLType = Return_OleDBType(datatable.Columns[i].DataType.ToString());
				    updatecommand.Parameters.Add(new OleDbParameter(datatable.Columns[i].Caption, colSQLType,0, datatable.Columns[i].Caption));
				}

				updateString2 += "([" + datatable.Columns[i].Caption + "] = ? OR ? IS NULL AND [" + datatable.Columns[i].Caption + "] IS NULL)"; 
						
				if (i!= datatable.Columns.Count-1) updateString2 +=" AND ";

				else {updateString1 += " WHERE (" + updateString2 + ")";}
			}

			for (int i=0;i< datatable.Columns.Count;i++)
			{
				OleDbType colSQLType = Return_OleDBType(datatable.Columns[i].DataType.ToString());
				updatecommand.Parameters.Add(new OleDbParameter("Original_" + datatable.Columns[i].Caption, colSQLType, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(10)), ((System.Byte)(10)), datatable.Columns[i].Caption, System.Data.DataRowVersion.Original, null));
				updatecommand.Parameters.Add(new OleDbParameter("Original_" + datatable.Columns[i].Caption + "1", colSQLType, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(10)), ((System.Byte)(10)), datatable.Columns[i].Caption, System.Data.DataRowVersion.Original, null));
			}

			updatecommand.CommandText = updateString1;
		}

		private OleDbType Return_OleDBType(string SystemType)
		{
			switch(SystemType)
			{

				case "System.Boolean":
					return OleDbType.Boolean;
				case "System.Byte":
					return OleDbType.UnsignedTinyInt;
				case "System.Byte[]":
					return OleDbType.VarBinary;
				case "System.Binary":
					return OleDbType.VarBinary;
				case "System.DateTime":
					return OleDbType.DBTimeStamp;
				case "System.Decimal":
					return OleDbType.Decimal;
				case "System.Double":
					return OleDbType.Double;
				case "System.Single":
					return OleDbType.Single;
				case "System.Guid":
					return OleDbType.Guid;
				case "System.Int16":
					return OleDbType.SmallInt;
				case "System.Int32":
					return OleDbType.Integer;
				case "System.Int64":
					return OleDbType.BigInt;
				case "System.Object":
					return OleDbType.Variant;
				case "System.String":
					return OleDbType.VarWChar;
				case "System.UInt16":
					return OleDbType.UnsignedSmallInt;
				case "System.UInt32":
					return OleDbType.UnsignedInt;
				case "System.UInt64":
					return OleDbType.UnsignedBigInt;
				case "System.AnsiString":
					return OleDbType.VarChar;
				case "System.Currency":
					return OleDbType.Currency;
				case "System.Date":
					return OleDbType.DBDate;
				case "System.SByte":
					return OleDbType.TinyInt;
				case "System.Time":
					return OleDbType.DBTime;
				case "VarNumeric":
					return OleDbType.VarNumeric;
				default:
					return OleDbType.Variant;
			}
		}
		#endregion

		#region "make_Delete_Command"
		//Here the command string and the parameters are
		//assigned for the delete command
		private void make_Delete_Command(DataTable datatable, OleDbCommand deletecommand)
		{
			string deleteString = "DELETE FROM [" + comboTables.Text + "] WHERE ";

			for (int i=0;i< datatable.Columns.Count;i++)
			{
				deleteString +="( [" + datatable.Columns[i].Caption + "] = ? OR ? IS NULL AND [" + datatable.Columns[i].Caption + "] IS NULL )"; 
						
				if (i!= datatable.Columns.Count-1){deleteString +=" AND ";}
				
				deletecommand.Parameters.Add(new OleDbParameter("Original_" + datatable.Columns[i].Caption, System.Data.OleDb.OleDbType.Variant, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(10)), ((System.Byte)(10)), datatable.Columns[i].Caption, System.Data.DataRowVersion.Original, null));
				deletecommand.Parameters.Add(new OleDbParameter("Original_" + datatable.Columns[i].Caption + "1", System.Data.OleDb.OleDbType.Variant, 0, System.Data.ParameterDirection.Input, false, ((System.Byte)(10)), ((System.Byte)(10)), datatable.Columns[i].Caption, System.Data.DataRowVersion.Original, null));
			}

			deletecommand.CommandText = deleteString;
		}
		#endregion

		#region "Check_If_Data_Changed"
		// Here we check if the data has changed
		private bool Check_If_Data_Changed()
		{
			try
			{	
				// Create a new dataset to hold the changes that have been made to the main dataset.
				DataSet objDataSetChanges = new DataSet();
				// Stop any current edits.
				this.BindingContext[accessDataSet,comboTables.Text].EndCurrentEdit();
				// Get the changes that have been made to the main dataset.
				objDataSetChanges = ((DataSet)(accessDataSet.GetChanges()));
				// Check to see if any changes have been made.
				if ((objDataSetChanges != null)) return true;
				else return false;
			}
			catch
			{return false;}
		}
		#endregion

		#region "arrange_Controls"
		//Here the control positions and other control paramters are
		//set based on the data loaded
		private void arrange_Controls(int startingPos, int TabIndex_start)
		{
			btnAdd.TabIndex = TabIndex_start+1;
			btnCancel.TabIndex = TabIndex_start+2;
			//btnDelete.TabIndex = TabIndex_start+3;
			btnUpdate.TabIndex = TabIndex_start+5;
            Findbtn.TabIndex = TabIndex_start+6;
            textBox2.TabIndex = TabIndex_start + 7;

			comboTables.Top = startingPos + 42;
			comboTables.Left = 10;

			//lblcombo.Top = startingPos + 45;
			//lblcombo.Left = 10;

			btnNavNext.Top = startingPos + 70;
            btnNavNext.Left = 174;  // 138;

			lblNavLocation.Top = startingPos + 70;
			lblNavLocation.Left = 60; //42;

			btnNavPrev.Top = startingPos + 70;
            btnNavPrev.Left = 10;


            
            //textBox2.Top = startingPos + 70;
            //textBox2.Left = 492;
            //this.textBox2.Margin.Top = 70;
            //this.textBox2.Margin.Left = 400;
            //GroupBox groupBox1 = new GroupBox();
            //TextBox textBox3 = new TextBox();
            this.textBox2.Location = new Point(30, 15);
            this.textBox2.Size = new Size(75, 21);
            this.textBox2.Top = startingPos + 70;
            this.textBox2.Left = 492;

            //groupBox1.Controls.Add(textBox3);
            // Set the Text and Dock properties of the GroupBox.
            //groupBox1.Text = "MyGroupBox";
            //groupBox1.Dock = DockStyle.Bottom;
            //groupBox1.Top = startingPos + 40;
            //groupBox1.Left = 300;

            // Disable the GroupBox (which disables all its child controls)
            //groupBox1.Enabled = false;

            // Add the Groupbox to the form. 
            //this.Controls.Add(textBox3);
            this.Controls.Add(this.textBox2);
            //textBox2.Visible = true;
            //textBox3.Visible = true;
            

            



            Findbtn.Top = startingPos + 70;
            Findbtn.Left = 573;

			btnCancel.Top = startingPos + 42;
			btnCancel.Left = 573;

			//btnDelete.Top = startingPos + 42;
			//btnDelete.Left = 654;

            btnUpdate.Top = startingPos + 42;
            btnUpdate.Left = 654;

			btnAdd.Top = startingPos + 42;
			btnAdd.Left = 492;

			DataLoaded = true;

		    //NiceMenu.myModifyNiceMenu[0].MenuItems[2].Enabled = true;
    		//NiceMenu.myModifyNiceMenu[0].MenuItems[3].Enabled = true;
			//NiceMenu.myModifyNiceMenu[1].MenuItems[0].Enabled = true;

			btnAdd.Visible = true;
			btnCancel.Visible = true;
			//btnDelete.Visible = true;
			btnUpdate.Visible = true;
			btnNavPrev.Visible = true;
			btnNavNext.Visible = true;
			comboTables.Visible = true;
			lblNavLocation.Visible = true;
            //groupBox1.Visible = true;
            //textBox2.Visible = true;
            Findbtn.Visible = true;
            
            

			GPostion = startingPos + 170;
			dGrid.Height= this.Height-GPostion;
		}
		#endregion
		
		#region "DataHighlighted Position Change"
		//update the lblnavlocation text
		private void DataSet_PositionChanged()
		{
            try
            {
                if (Check_If_Data_Changed() == true)
                {
                    try
                    {
                        this.UpdateDataSet();
                        this.accessDataAdapter.Fill(accessDataSet.Tables[comboTables.Text]);
                        this.dGrid.Invalidate();
                        this.dGrid.Refresh();
                    }
                    catch (System.Data.OleDb.OleDbException eUpdate)
                    {
                        // Add your error handling code here.
                        // Display error message, if any.
                        System.Windows.Forms.MessageBox.Show(eUpdate.Message);
                    }

                   /*
                    DialogResult r = MessageBox.Show("The data was modified, do you want to save (Yes) or override (No) changes and proceed?", "Recent updates are pending", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
                    if (r == DialogResult.Yes)
                    {
                        this.UpdateDataSet();
                    }
                    else if (r == DialogResult.No)
                    {
                        //LoadData("SELECT * FROM [" + comboTables.Text + "]");     

                        this.accessDataAdapter.Fill(accessDataSet.Tables[comboTables.Text]);
                        this.dGrid.Invalidate();
                        this.dGrid.Refresh();  
                    }
                    //else return;
                    */
                }
               
                this.lblNavLocation.Text = ((((this.BindingContext[accessDataSet, comboTables.Text].Position + 1)).ToString() + " of  ")
                        + this.BindingContext[accessDataSet, comboTables.Text].Count.ToString());                     
            }
			catch (System.Data.OleDb.OleDbException eUpdate) 
			{
				// Add your error handling code here.
				// Display error message, if any.
				System.Windows.Forms.MessageBox.Show(eUpdate.Message);
			}
		}
		#endregion

		#region "Update Records in database"
		//This routine handles and perfroms the update
		//procedure to save changes in the main source
        /*
		public void UpdateDataSet()
		{
			// Stop any current edits.
			//this.BindingContext[accessDataSet,comboTables.Text].EndCurrentEdit();
            this.BindingContext[accessDataSet, comboTables.Text].SuspendBinding();
            
			// Check to see if any changes have been made.
            if (accessDataSet.HasChanges(DataRowState.Modified)) 
			{
                UpdateDataSource(accessDataSet);
				if (this.accessConnection.State.ToString()!="Closed") this.accessConnection.Close();
			}
            this.BindingContext[accessDataSet, comboTables.Text].ResumeBinding();
		}
        */

        public void UpdateDataSet()
		{
			// Create a new dataset to hold the changes that have been made to the main dataset.
			DataSet objDataSetChanges = new DataSet();
			// Stop any current edits.
			this.BindingContext[accessDataSet,comboTables.Text].EndCurrentEdit();
			
            // Get the changes that have been made to the main dataset.
			objDataSetChanges = ((DataSet)(accessDataSet.GetChanges()));
            
			// Check to see if any changes have been made.
			if ((objDataSetChanges != null)) 
            //if (accessDataSet.HasChanges(DataRowState.Modified))
            {
                UpdateDataSource(objDataSetChanges);
                if (this.accessConnection.State.ToString() != "Closed") this.accessConnection.Close();
            }
            else
            {
                MessageBox.Show("The data was compared to master database and no differences were detected. No updates were found to be saved.", "File is not updated");
            }
		}
         

        public void UpdateDataSource(DataSet ChangedRows)
		{
			// The data source only needs to be updated if there are changes pending.
			if ((ChangedRows != null)) 
			{
				// Open the connection.
				this.accessConnection.Close();
                string executable = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string path = (System.IO.Path.GetDirectoryName(executable));
                AppDomain.CurrentDomain.SetData("DataDirectory", path);
                store_accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=|DataDirectory|\Database63_fe.mdb";
                
                try
                {
                    //this.accessConnection.Open();
                    store_accessConnection.Open();
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Failed to connect to data source", "Warning");
                    MessageBox.Show("Connecting to data source from this system... Please click OK", "Warning");
                }
                 
				// Attempt to update the data source.
				try
				{
					accessDataAdapter.Update(ChangedRows);
					accessDataSet.AcceptChanges();
				}
				//Catch all the erros and report them
				catch(System.Data.DBConcurrencyException e)
				{
					accessDataSet.AcceptChanges();
					MessageBox.Show("Unfortunately this data file cannot be saved. " +
						"Concurrency violation (please exit, " +
						"all data are rejected, renter your data).","Cannot Update File");
				}
				finally 
				{
					//this.accessConnection.Close();
                    store_accessConnection.Close();
				}
			}
		}
		#endregion

		#region "This Form's All Button Click Events"
		//This section holds all the click events for all
		//the controls on the form

		//Navigate Previous Routine
		private void btnNavPrev_Click(object sender, System.EventArgs e)
		{
			this.BindingContext[accessDataSet,comboTables.Text].Position = (this.BindingContext[accessDataSet,comboTables.Text].Position - 1);
			//this.DataSet_PositionChanged();
		}

		//Navigate Next Routine
		private void btnNavNext_Click(object sender, System.EventArgs e)
		{
			this.BindingContext[accessDataSet,comboTables.Text].Position = (this.BindingContext[accessDataSet,comboTables.Text].Position + 1);
			//this.DataSet_PositionChanged();
		}

		//Add Data Routine
		private void btnAdd_Click(object sender, System.EventArgs e)
		{
            DialogResult R = MessageBox.Show("Warning: A new record will be created. Records can not be deleted. Yes - Add, No - Cancel", "Adding New Record", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (R == DialogResult.Yes) 
            {
                try
                {
                    // Clear out the current edits
                    this.BindingContext[accessDataSet, comboTables.Text].EndCurrentEdit();
                    this.BindingContext[accessDataSet, comboTables.Text].AddNew();
                }
                catch (System.Exception eEndEdit)
                {
                    System.Windows.Forms.MessageBox.Show(eEndEdit.Message);
                }

                this.BindingContext[accessDataSet, comboTables.Text].EndCurrentEdit();
            }
		}


		//Cancel Current Data Entry Routine
		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.BindingContext[accessDataSet,comboTables.Text].CancelCurrentEdit();
			//this.DataSet_PositionChanged();
		}

		//Update and Save Data Changes in the source file
		private void btnUpdate_Click(object sender, System.EventArgs e)
		{
            Cursor.Current = Cursors.WaitCursor;
			try 
			{
                // Update the datasource and reload.
                this.UpdateDataSet();
			}
			catch (System.Data.OleDb.OleDbException eUpdate) 
			{
				// Add your error handling code here.
				// Display error message, if any.
				System.Windows.Forms.MessageBox.Show(eUpdate.Message);
			}

 			//this.DataSet_PositionChanged();
            Cursor.Current = Cursors.Arrow;
		}


        //Cancel Current Data Entry Routine
        private void btnWordDoc_Click(object sender, System.EventArgs e)
        {
            Form create_word_doc_from_template = new BuildDocFrm(this, accessDataSet, dGrid, comboTables);
            create_word_doc_from_template.Show();
        }


		//AutoNumber or Automatic Current Date Entry Routine
		private void cMenuClick(object sender, System.EventArgs e)
		{
			MenuItem tempItem = (MenuItem)sender;
			tempItem.Checked = !tempItem.Checked;
		}

		#endregion

		#region "removeMadeControls"
		//Here all the made controls are removed
		private void removeMadeControls()
		{
			for (int i=0;i<this.Controls.Count;i++) 
			{
				if ((this.Controls[i].GetType().Name == "Label") &&
					(this.Controls[i].Name != "lblNavLocation"))
				{
					this.Controls[i].Dispose();
					i=0;
				}
				else if (this.Controls[i].GetType().Name == "TextBox")
				{
					this.Controls[i].Dispose();
					i=0;
				}
			}
		}
		#endregion

		#region "comboTables_SelectedValueChanged"		
		//ComboTables Value Changed Routine
		private void comboTables_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (ComboBoxText == comboTables.Text) return;
			else ComboBoxText = comboTables.Text;

			removeMadeControls();
			LoadData("SELECT * FROM [" + comboTables.Text + "]");
		}
		#endregion

		#region "Form Closing and Resizing Event"


        //You can override OnFormClosing to do this. Just be careful you don't do 
        //anything too unexpected, as clicking the 'X' to close is a well understood behavior.
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            mainform = new MainFrm();
            base.OnFormClosing(e);

            if (e.CloseReason == CloseReason.WindowsShutDown
                || e.CloseReason == CloseReason.TaskManagerClosing) return;

            if (Check_If_Data_Changed() == true)
            {
                DialogResult r = MessageBox.Show("The database file changed, are you sure you want exit without saving?", "Exit Without Save", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (r == DialogResult.Yes)
                {
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
            }
            else
            {
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
            mainform = null;
            this.Cursor = Cursors.Arrow;
        }
       
		private void MainFrm_Resize(object sender, System.EventArgs e)
		{
			dGrid.Height= this.Height-GPostion;
		}
		#endregion

		#region "Data Grid's Event"

		//Show AutoNumber Incrementation or Automatic Date
		//Insertion menu creation routine 
		private void dGrid_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			//Runs for right mouse down
			DataGrid myGrid = (DataGrid) sender;
			System.Windows.Forms.DataGrid.HitTestInfo hti;

			hti = myGrid.HitTest(e.X, e.Y);
            
			switch (hti.Type) 
			{
				case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader :
					if ((colType[hti.Column]=="System.Int32")|| (colType[hti.Column]=="System.DateTime"))
					{
						if (e.Button != System.Windows.Forms.MouseButtons.Right) return;
						AutoMenu[hti.Column].Show(dGrid,new Point(e.X,e.Y));
					}
					break;
			}
		}

		//The Grid's update routine if any of the menue's are
		//checked
		private void dGrid_CurrentCellChanged(object sender, System.EventArgs e)
		{      
			    this.DataSet_PositionChanged();
			    string[] getRow = lblNavLocation.Text.Split(' ');
                try
                {
                    if (getRow[0]!=getRow[3]) return;
                }
                catch (System.IndexOutOfRangeException)
                {
                    //return;
                }
			    DataGrid myGrid = (DataGrid) sender;
			    int row = Convert.ToInt32(getRow[0]);
    					
			    for (int i=0;i<accessDataSet.Tables[comboTables.Text].Columns.Count;i++)
			    {
				    if ((colType[i]=="System.Int32")&&(AutoMenu[i].MenuItems[0].Checked==true))
				    {
					    if (row > 1)
					    {
						    try 
						    {
							    UpdateTextBoxes(i, Convert.ToString((int)dGrid[dGrid.CurrentCell.RowNumber-1, i]+1));
						    }
						    catch(System.InvalidCastException)
						    {
							    return;
						    }
					    }
				    }
				    if ((colType[i]=="System.DateTime")&&(AutoMenu[i].MenuItems[0].Checked==true))
				    {
					    {
                            UpdateTextBoxes(i, System.DateTime.Now.Date.ToString());
					    }
				    }
			    }
            
		}
		
		//Menu Updates done in the textbox controls
		//after editing here
		private void UpdateTextBoxes(int Col, string newValue)
		{
			for (int i=0; i< this.Controls.Count;i++)
			{
				if (this.Controls[i].Name ==("cText"+Col.ToString()))
				{
					if (this.Controls[i].Text!="") return;
					this.Controls[i].Text = newValue;
				}
			}
		}
		#endregion


        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            ////MessageBox.Show(" 1 does not exists in the list.\r\n ");
            //Form helpform = new Form1();
            //helpform.Show();
            
            //Help.ShowHelp(this, "file://C:\\Law\\EasyClientMaster\\Easy Master Client Help.chm");

            string fbPath = Application.StartupPath;
            string fname = "Easy Master Client Help.chm";
            string filename = fbPath + @"\" + fname;
            //Help.ShowHelp(this, filename);

            System.Diagnostics.Process.Start(filename);            
        }

       

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            ColorDialog diag = new ColorDialog();
            //diag.ShowDialog();
            
            if (diag.ShowDialog() == DialogResult.OK)
            {
                BackColor = diag.Color;
            }
        }


        private void View_Same_Test_in_Bigger_Window(object sender, EventArgs e)
        {
            Exception X = new Exception();
            RichTextBox T = (RichTextBox)sender;

            Form display_textField_in_window = new ViewTextFieldInBigWindow(T.Text, T);
            try
            {
                display_textField_in_window.Show();
            }
            catch (Exception)
            {
                try
                {
                    MessageBox.Show("Warning: Attempt to modify existing data. Please Enter a different value.");
                }
                catch (Exception) { }
            }
        }

        //Open_Outlook_email
        private void Open_Outlook_email(object sender, EventArgs e)
        {
            Exception X = new Exception();
            RichTextBox T = (RichTextBox)sender;
            int clmn_clientName = 2;
            int clmn_lawyerName = 5;
            int row_number = this.BindingContext[accessDataSet, comboTables.Text].Position;
            string client_name = this.accessDataSet.Tables[comboTables.Text].Rows[row_number][clmn_clientName].ToString();
            string lawyer_name = this.accessDataSet.Tables[comboTables.Text].Rows[row_number][clmn_lawyerName].ToString();
            //int number_of_rows_in_table = this.accessDataSet.Tables[comboTables.Text].Rows.Count;
            //int number_of_clmns_in_table = this.accessDataSet.Tables[comboTables.Text].Columns.Count;

            try
            {
                Outlook.Application oApp = null;
                Outlook.MailItem oMsg = null;
                Outlook.Inspector oAddSig = null;



                // Create the Outlook application.
                oApp = new Outlook.Application();
                // Create a new mail item.
                oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                oAddSig = oMsg.GetInspector;   //Needed for Outlook 2007
                // Set HTMLBody. 
                //add the body of the email
                oMsg.Body = "Dear " + client_name + "\n\n    \rSincerely, " + "\n\n" + lawyer_name;
                //Add an attachment.
                String sDisplayName = "MyAttachment";
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //now attached the file
                //Outlook.Attachment oAttach = oMsg.Attachments.Add(@"C:\\fileName.jpg", iAttachType, iPosition, sDisplayName);
                //Subject line
                oMsg.Subject = "";
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(T.Text);
                oRecip.Resolve();
                // Send.
                try
                {
                    oMsg.Display(true);//.Send();
                    
                }
                catch (Exception) { MessageBox.Show("Warning: Can not open Outlook."); }
                
                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }//end of try block
            catch (Exception)
            {
                try
                {
                    MessageBox.Show("Warning: Can not open Outlook. Please verify that e-mail address is specified correctly in this file.");
                }
                catch (Exception) { }
            }
        }


        //Function/method to validate fields format
        private void HasSpecialChars(object sender, EventArgs e)
        {
            Exception X = new Exception();
            RichTextBox T = (RichTextBox)sender;
            string current_cell_value_dgrid = "";
            int current_row_number = (int)dGrid.CurrentCell.RowNumber;

            current_cell_value_dgrid = this.accessDataSet.Tables[comboTables.Text].Rows[current_row_number][T.DataBindings["Text"].FormatString].ToString();

            try
            {
                if (T.DataBindings["Text"].FormatString == "First Name" || T.DataBindings["Text"].FormatString == "Last Name")
                {
                    if (!System.Text.RegularExpressions.Regex.IsMatch(T.Text, "^[a-zA-Z -]*$"))
                    {
                        MessageBox.Show("Allowed characters in the " + T.DataBindings["Text"].FormatString + ": Text only ", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        //T.Text = current_cell_value_dgrid;
                    }
                }

                else if (T.DataBindings["Text"].FormatString == "E-mail")
                {
                    if (T.Text != "" && (T.Text.IndexOf("@", StringComparison.OrdinalIgnoreCase) == -1 || T.Text.IndexOf(".", StringComparison.OrdinalIgnoreCase) == -1))
                    {
                        MessageBox.Show("Incorrect format of the e-mail address. Please re-enter", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        //T.Text = current_cell_value_dgrid;
                    }
                }
                else if (T.DataBindings["Text"].FormatString == "Phone Number" || T.DataBindings["Text"].FormatString == "Call Number")
                {
                    //if (T.Text.IndexOf("-", StringComparison.OrdinalIgnoreCase) == -1 || T.Text.IndexOf(".", StringComparison.OrdinalIgnoreCase) == -1)
                    if (!System.Text.RegularExpressions.Regex.IsMatch(T.Text, "^[0-9-()]*$"))
                    {
                        MessageBox.Show("Incorrect format of the phone number. Please re-enter", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        //T.Text = current_cell_value_dgrid;
                    }
                }
            }
            catch (Exception)
            {
                try
                {
                    MessageBox.Show("Attempt to modify existing data. Please Enter a different value.", "WARNING!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (Exception) { }
            }
        }


        #region "Methods for Find Button feature"
        private void TextControl_MouseClickTextBox(object sender, EventArgs e)
        {
            TextBox T = (TextBox)sender;

            if( T.Text == "search field" &&  T.ForeColor == Color.SkyBlue)
            {
            	T.Text = T.Text.Replace("search field", "");
            	T.ForeColor = Color.Black;
            }
        }
       
        private void TextControl_Modified(object sender, EventArgs e)
        {
            TextBox T = (TextBox)sender;
            
            if( T.Modified == false) return;
            
            if( T.TextLength > 0 && T.ForeColor == Color.SkyBlue)
            {
                T.Text = remeberLastKey;
            	T.ForeColor = Color.Black;
                T.SelectionStart = T.Text.Length;
            } 
            else if( T.TextLength > 0 && T.ForeColor != Color.SkyBlue)
            {
            	T.ForeColor = Color.Black;
            }
            else if( T.TextLength == 0 )
            {
            	T.ForeColor = Color.SkyBlue;
            	T.Text = "search field";
            	T.SelectionStart = 0;
            }

            //reset the Modified property back to default true value for the next change in the textbox to be triggered
            T.Modified = false;
        }

        private void CheckKeys(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            TextBox T = (TextBox)sender;

            if (e.KeyChar == (char)13)  //if key = ENTER
            {
                this.Findbtn_Click(T, e);
            }
            else if (e.KeyChar == (char)08)
                remeberLastKey = "";
            else
                remeberLastKey = e.KeyChar.ToString();
        }

        private void TextControl_MouseLeave(object sender, EventArgs e)
        {
            TextBox T = (TextBox)sender;

            T.ForeColor = Color.SkyBlue;
            T.Text = "search field";
            T.SelectionStart = 0;
        }
        #endregion


        private void Findbtn_Click(object sender, EventArgs e)
        {           
            if (String.IsNullOrEmpty(this.textBox2.Text) || this.textBox2.ForeColor == Color.SkyBlue) return;

            String searchValue = this.textBox2.Text;
            
            int rowIndex = -1;

            dGrid.UnSelect(dGrid.CurrentRowIndex);
            Cursor.Current = Cursors.WaitCursor;

            int i = this.BindingContext[accessDataSet, comboTables.Text].Position;
            int number_of_rows_in_table = this.accessDataSet.Tables[comboTables.Text].Rows.Count;
            int number_of_clmns_in_table = this.accessDataSet.Tables[comboTables.Text].Columns.Count;

            for (int var_count_rows_in_loop = 0; var_count_rows_in_loop < number_of_rows_in_table; var_count_rows_in_loop++)
            {
                i++;					   //increment starting at the current cursor position in a table		
                if (i + 1 > number_of_rows_in_table)
                {
                    //MessageBox.Show("End of Search: " + var_count_rows_in_loop.ToString());
                    i = i - number_of_rows_in_table;  //move to the top of table to continue loop
                }


                for (int clmn = 0; clmn < number_of_clmns_in_table; clmn++)
                {
                   
                    //if (this.accessDataSet.Tables[comboTables.Text].Rows[i][clmn].ToString().IndexOf(searchValue, StringComparison.OrdinalIgnoreCase) != -1) //This is wrong. This way it finds the record in the back end Table on the network, but not in the datagrid in UI
                    if (dGrid[i,clmn].ToString().IndexOf(searchValue, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        rowIndex = i;

                        //Move cursor to a new position in datagrid
                        //and display a new data in text boxes
                        this.BindingContext[accessDataSet, comboTables.Text].Position = rowIndex;
                        //this.DataSet_PositionChanged();

                        //dGrid.CurrentCell = new DataGridCell(i, clmn);  //select/highlight current cell in the datagrid table
                        var_count_rows_in_loop = this.accessDataSet.Tables[comboTables.Text].Rows.Count;  //stop cout main FOR loop cycle and break

                        Cursor.Current = Cursors.Default;  // Back to normal 
            			dGrid.Select(dGrid.CurrentRowIndex);  //select entire row in the datagrid table where the item was found
                    
            			break;
                        //return;
                    }
                }
                
            }
           
            //this.BindingContext[accessDataSet, comboTables.Text].Position = rowIndex;
            //this.DataSet_PositionChanged();
            
            if (rowIndex == -1)
                MessageBox.Show("No matches were found.");           
        }

        //DB Reload button feature - refresh data in the field on the form
        private void btn_Refresh(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //removeMadeControls();

            if (DataLoaded == false)
            {
                MessageBox.Show("The database file has to be loaded.");
            }
            else if (Check_If_Data_Changed() == true)
            {
               DialogResult r = MessageBox.Show("The database file changed, are you sure you want to proceed and override your unsaved edits?", "Go on without saving", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
               if (r == DialogResult.Yes)
               {
                   //LoadData("SELECT * FROM [" + comboTables.Text + "]");     

                   this.accessDataAdapter.Fill(accessDataSet.Tables[comboTables.Text]);
                   this.dGrid.Invalidate();
                   this.dGrid.Refresh();
               }
            }
            else
            {
                DialogResult r = MessageBox.Show("Refresh Screen action will override unsaved changes if any, are you sure you want to proceed?", "Load data from database", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (r == DialogResult.Yes)
                {
                    //LoadData("SELECT * FROM [" + comboTables.Text + "]");     

                    this.accessDataAdapter.Fill(accessDataSet.Tables[comboTables.Text]);
                    this.dGrid.Invalidate();
                    this.dGrid.Refresh();
                }
            }
            this.Cursor = Cursors.Arrow;
        }
 
        public void menuItem6_BuildDocMailMergeClick(object sender, EventArgs e)
        {
            BuildDocMailMerge mail_merge_doc = new BuildDocMailMerge(accessDataSet, dGrid, comboTables);
            mail_merge_doc.button1_Click(sender, e);
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Save the changes to the source database the mdb file
            btnUpdate.Focus();
            DialogResult R = MessageBox.Show("Are you sure you want to save? Changes will be permenant.", "Save Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (R == DialogResult.Yes) btnUpdate_Click(sender, e);
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            mainform = new MainFrm();

            if (Check_If_Data_Changed() == true)
            {
                DialogResult r = MessageBox.Show("The database file changed, are you sure you want exit without saving?", "Exit Without Save", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (r == DialogResult.Yes)
                {
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
                  //  e.Cancel = true;
            }
            else
            {
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
            mainform = null;
            Cursor.Current = Cursors.Arrow;
        }

        private void mailMergeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DataLoaded == false)
            {
                MessageBox.Show("The database file has to be loaded.");
            }
            else if (Check_If_Data_Changed() == true)
            {
                DialogResult r = MessageBox.Show("The database file changed, are you sure you want to proceed?", "Go on without saving", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (r == DialogResult.Yes) menuItem6_BuildDocMailMergeClick(sender, e);
            }
            else
            {
                menuItem6_BuildDocMailMergeClick(sender, e);
            }
        }
 
    }
}