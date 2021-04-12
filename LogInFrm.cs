using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Diagnostics;


namespace DataEasy
{
    public partial class LogInFrm : Form
    {
        #region "Private Declarations"
        private OleDbConnection accessConnection = new OleDbConnection();
        #endregion

        [STAThread]
        static void Main()
        {
            Application.Run(new LogInFrm());
        }

        public LogInFrm()
        {
            InitializeComponent(); 
            autoOpenDatabase();
            this.textBox1.Focus();
            this.textBox1.Select();
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
                Cursor.Current = Cursors.WaitCursor;
                this.Text = "MasterClient: Loading Data Please Wait...";

                ////The connection parameters
                string executable = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string path = (System.IO.Path.GetDirectoryName(executable));
                AppDomain.CurrentDomain.SetData("DataDirectory", path);
                this.accessConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;data source=|DataDirectory|\Database63_fe.mdb";

                // Open the connection.
                if (this.accessConnection.State.ToString() != "Open")
                {
                    try
                    {
                        //MessageBox.Show("Open Database");
                        this.accessConnection.Open();
                        Cursor.Current = Cursors.Default;  // Back to normal
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Connecting to data source from this system... Please click OK", "Warning");
                    }
                }
            }
            catch (System.Data.OleDb.OleDbException eFillDataSet)
            {
                MessageBox.Show("Exception loading the database");
                throw eFillDataSet;
            }
            //Return the cursor and form's caption to their normal state
            this.Cursor = Cursors.Arrow;
            this.Text = "Easy Master Client - Login";
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                string cmdText = "select Count(*) from Login where Username=? and [Password]=?";
                using (OleDbCommand cmd = new OleDbCommand(cmdText, accessConnection))
                {
                    cmd.Parameters.AddWithValue("@p1", textBox1.Text);
                    cmd.Parameters.AddWithValue("@p2", textBox2.Text);
                    int result = (int)cmd.ExecuteScalar();
                    if (result > 0)
                    {
                        //MessageBox.Show("Login Successful");

                        if (this.accessConnection.State.ToString() == "Open")
                        {
                            //MessageBox.Show("Close Database");
                            this.accessConnection.Close();
                        }

                        //MainFrm mainform = new MainFrm();
                        Startup startupform = new Startup();
                        startupform.Show();

                        this.Hide();
                        //this.Close();
                        
                    }
                    else
                    {
                        MessageBox.Show("Invalid Credentials, Please Re-Enter");
                    }
                    Cursor.Current = Cursors.Default;  // Back to normal
                }
            }
            catch (Exception exl)
            {
                MessageBox.Show("Faild to connect to database!\nConnect to the server where the database is located and try again.\nYou may need to enter username and password.", "Error");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void CheckKeys(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            TextBox T = (TextBox)sender;

            if (e.KeyCode == Keys.Enter)  //if key = ENTER
            {
                this.button1_Click(T, e);
            }
            else if (e.KeyCode == Keys.F4) //if key = F4
            {
                if (this.accessConnection.State.ToString() == "Open")
                {
                    this.accessConnection.Close();
                }
                Startup startupform = new Startup();
                startupform.Show();
                this.Hide();
            }
        }
    }
}
