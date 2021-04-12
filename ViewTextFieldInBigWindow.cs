using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DataEasy
{
    public partial class ViewTextFieldInBigWindow : Form
    {
        //private bool whait_for_OKbtn_clicked = false;
        //private static string name = "";
        //public static string text_edit = "";
        private string store_text_edit = "";
        private RichTextBox T;

        public ViewTextFieldInBigWindow(string text_from_mnform, RichTextBox textox_object)
        {
            InitializeComponent();
            T = textox_object;

            textBox1.Text   = text_from_mnform;     //local textbox field gets value passed from mnForm          
            store_text_edit = text_from_mnform;     //backup initial text value
        }

        public void okButton_Click(object sender, EventArgs e)
        {
            if (store_text_edit != textBox1.Text)
            {
                T.Text = textBox1.Text;     //pass value back to mnForm
            }
            this.Close();
        }
    }
}