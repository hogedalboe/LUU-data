using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LUU_data
{
    public partial class Form5 : Form
    {
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }

        public Form5()
        {
            InitializeComponent();

            // [CL:2]
            string file_latestOutlookReading = Directory.GetCurrentDirectory() + @"\data\settings\latest-outlook-reading.txt";
            string latestFrom = "";
            string latestTo = "";
            string latestDate = "";
            string latestInit = "";

            Encoding danish = Encoding.GetEncoding(1252);

            DateTime dtToday = DateTime.Today;

            // Give the user info about the latest outlook reading [CL:2]
            try
            {
                string[] arrLatestReading = File.ReadAllLines(file_latestOutlookReading, danish);
                latestFrom = arrLatestReading[0];
                latestTo = arrLatestReading[1];
                latestDate = arrLatestReading[2];
                latestInit = arrLatestReading[3];

                label8.Text = latestFrom;
                label9.Text = latestTo;
                label10.Text = latestDate;
                label11.Text = latestInit;
            }
            catch
            {
                label8.Text = "";
                label9.Text = "";
                label10.Text = "";
                label11.Text = "";
            }

            // Set the start date [CL:2]
            if (latestTo != "")
            {
                try
                {
                    DateTime dtLatestTo = DateTime.ParseExact(latestTo, "dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    DateTime plusMonth = dtLatestTo.AddMonths(1);

                    dateTimePicker1.Value = dtLatestTo; // Start date = To date of latest reading

                    dateTimePicker2.Value = plusMonth;
                    /*
                    // Set either a month from dtLatestTo or today
                    if (plusMonth > dtToday)
                    {
                        dateTimePicker2.Value = dtToday;
                    }
                    else
                    {
                        dateTimePicker2.Value = plusMonth;
                    }
                    */
                }
                catch
                {
                    setDefaultInterval();
                }
            }
            else
            {
                setDefaultInterval();
            }

            // Set default interval [CL:2]
            void setDefaultInterval()
            {
                dateTimePicker1.Value = dtToday.AddMonths(-1);
                dateTimePicker2.Value = dtToday;
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (dateTimePicker2.Value >= dateTimePicker1.Value) // [CL:4]
            {
                // OK
                this.DialogResult = DialogResult.OK;

                this.startDate = dateTimePicker1.Value;
                this.endDate = dateTimePicker2.Value;

                this.Close();
            }
            else
            {
                MessageBox.Show("Den afgrænsede periode er ikke kronologisk.");
            }

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            // Abort
            this.DialogResult = DialogResult.Abort;
            this.Close();
        }
    }
}
