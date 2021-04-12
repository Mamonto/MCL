using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Collections.Generic;


namespace DataEasy
{
    public partial class ChartStats : Form
    {
        Dictionary<string, int> dict = new Dictionary<string, int>();
        string column_header;
        int counter = 0;

        //Series series;
        //ChartType = SeriesChartType.Pie;

        public ChartStats(Dictionary<string, int> dict_from_dgrid, string header_title)
        {
            InitializeComponent();
            dict = dict_from_dgrid;
            column_header = header_title;

            Form1_Load();
        }

        private void Form1_Load()
        {
            this.chart1.Series.Clear();
            this.chart1.Titles.Clear();

            // Set palette.
            this.chart1.Palette = ChartColorPalette.SeaGreen;
            
            // Set title.
            this.chart1.Titles.Add(column_header);

            // Add series.
            Series series = new Series();
            series.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;

            foreach (KeyValuePair<string, int> kv in dict)
            {
                // //MessageBox.Show(kv.Key.ToString());   //name
                // //MessageBox.Show(kv.Value.ToString()); //count

                // Add series.
                //seriesArray[counter] = kv.Key.ToString();
                //Series series = this.chart1.Series.Add(seriesArray[counter]);
                series = this.chart1.Series.Add(kv.Key.ToString());
                
                // Add point.
                //pointsArray[counter] = kv.Value;
                //series.Points.Add(pointsArray[counter]);
                series.Points.Add(kv.Value);
            }
        }

        // Load Pie
        private void Form2_Load()
        {
            this.chart1.Series.Clear();
            this.chart1.Titles.Clear();

            // Set palette.
            this.chart1.Palette = ChartColorPalette.SeaGreen;

            // Set title.
            this.chart1.Titles.Add(column_header);

            // Add series.
            Series series = this.chart1.Series.Add(column_header);

            // this.chart1.Series[0].Points.Clear();
            this.chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            foreach (string tagname in dict.Keys)
            {
                this.chart1.Series[0].Points.AddXY(tagname, dict[tagname]);
                //chart1.Series[0].IsValueShownAsLabel = true;
            }
        }

        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //this.chart1.SaveImage("C:\\MCL\\chart.png", ChartImageFormat.Png);
            SaveFileDialog dialog = new SaveFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.chart1.SaveImage(dialog.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }

        private void pieToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2_Load();
        }

        private void barToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form1_Load();
        }
    }
}
