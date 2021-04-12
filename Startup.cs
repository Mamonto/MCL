using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace DataEasy
{
    public partial class Startup : Form
    {
        private OleDbConnection accessConnection = new OleDbConnection();
        private string dataSourceFile = ""; //source file of the .mdb file
        MainFrm mainform = null;
        searchfrm sfrm = null;
        IncomingCallsFrm incomingcalls = null;

        public Startup()
        {
            InitializeComponent();
        }

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

                ////Change the mouse icon and caption 
                ////of the form to inform the user that
                ////the data is being loaded
                //Cursor.Current = Cursors.WaitCursor;
                this.Text = "MasterClient: Loading Data Please Wait...";

                //The connection parameters
                dataSourceFile = openFile.FileName;

                ////The connection parameters
                string executable = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string path = (System.IO.Path.GetDirectoryName(executable));
                AppDomain.CurrentDomain.SetData("DataDirectory", path);
                this.accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=|DataDirectory|\Database63_fe.mdb";

                // Open the connection.
                try
                {
                    if (this.accessConnection.State.ToString() != "Open")
                    {
                        //MessageBox.Show("Open Database");
                        this.accessConnection.Open();
                        //Cursor.Current = Cursors.Default;  // Back to normal
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Connecting to data source from this system... Please click OK", "Warning");
                }
            }
            catch (System.Data.OleDb.OleDbException eFillDataSet)
            {
                MessageBox.Show("Exception loading the database");
                throw eFillDataSet;
            }
            //Return the cursor and form's caption to their normal state
            //Cursor.Current = Cursors.Arrow;
            //this.Text = "Easy Master Client";
        }
        #endregion

        //Start Main Form
        private void button1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            //Create a new instance of the mnForm form
            if (mainform == null || mainform.Text == "")
            {
                mainform = new MainFrm();
                mainform.Show();
                this.Hide();
            }
            else if (CheckOpened(mainform.Text))
            {
                mainform.WindowState = FormWindowState.Normal;
                mainform.Focus();
            }
            else
                MessageBox.Show("ERROR - Can't open Main Form");

            mainform = null;
            Cursor.Current = Cursors.Arrow;
        }

        // Start Query/Search Form
        private void button2_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            //Create a new instance of the searchFrm form
            //Specifying the datasource and table to view
            if (sfrm == null || sfrm.Text == "")
            {
                autoOpenDatabase();
                sfrm = new searchfrm(this, dataSourceFile, "[Master Client List]", "Select * From [Master Client List]");
                sfrm.Show();
                this.Hide();
            }
            else if (CheckOpened(sfrm.Text))
            {
                sfrm.WindowState = FormWindowState.Normal;
                sfrm.Focus();
            }
            else
                MessageBox.Show("ERROR - Can't open Query Form");

            sfrm = null;
            Cursor.Current = Cursors.Arrow;
        }

        // Start Incoming Calls Form
        private void button3_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            //Create a new instance of the IncomingCallsFrm Form
            if (incomingcalls == null || incomingcalls.Text == "")
            {
                incomingcalls = new IncomingCallsFrm(this);
                incomingcalls.Show();
                this.Hide();
            }
            else if (CheckOpened(incomingcalls.Text))
            {
                incomingcalls.WindowState = FormWindowState.Normal;
                incomingcalls.Focus();
            }
            else
                MessageBox.Show("ERROR - Can't open Incoming Calls Form");

            incomingcalls = null;
            Cursor.Current = Cursors.Arrow;
        }

        //Check whether the form has been already open
        private bool CheckOpened(string name)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm.Text == name)
                {
                    return true;
                }
            }
            return false;
        }

        //Chage background image when mouse hover over button 1
        private void BackgroundImageBtn1_Change(object sender, EventArgs e)
        {
            this.toolTip1.Show("Client Management System", button1);
            this.BackgroundImage = global::EasyClientMaster.Properties.Resources.meeting_client_1098x587;
        }

        //Chage background image when mouse hover over button 2
        private void BackgroundImageBtn2_Change(object sender, EventArgs e)
        {
            this.toolTip1.Show("Queries, Statistics, History", button2);
            this.BackgroundImage = global::EasyClientMaster.Properties.Resources.statistics1_1000x624;
        }

        //Chage background image when mouse hover over button 3
        private void BackgroundImageBtn3_Change(object sender, EventArgs e)
        {
            this.toolTip1.Show("Register new incoming calls", button3);
            this.BackgroundImage = global::EasyClientMaster.Properties.Resources.phone_1098x732;
        }

        //Chage background image back to the default when mouse leaves a button
        private void BackgroundImage_ChangeBack(object sender, EventArgs e)
        {
            this.BackgroundImage = global::EasyClientMaster.Properties.Resources.businessSol_1098x841;
        }

        //You can override OnFormClosing to do this. Just be careful you don't do 
        //anything too unexpected, as clicking the 'X' to close is a well understood behavior.
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            if (e.CloseReason == CloseReason.WindowsShutDown
                || e.CloseReason == CloseReason.TaskManagerClosing) return;

            // Confirm user wants to close
            DialogResult r = MessageBox.Show("Are you sure you want to exit?", "Closing MCL", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (r == DialogResult.Yes)
            {
                e.Cancel = true;
                Application.ExitThread();
            }
            else
                e.Cancel = true;
        }

        // *** unused function  ***
        private void CheckKeys(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            TextBox T = (TextBox)sender;

            if (e.KeyCode == Keys.Enter)  //if key = ENTER
            {
                this.BackgroundImageBtn2_Change(T, e);
                this.button1_Click(T, e);
            }
        }

    }
}
