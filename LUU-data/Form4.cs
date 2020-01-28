using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LUU_data
{
    public partial class Form4 : Form
    {
        //////////////////////////////// Globals /////////////////////////////////////////

        string dir_summaries;
        string file_LUU;
        string file_summaries;

        string emailFrom;
        string emailReceived;
        string emailReceivedShortVersion; // [CL:6]
        string emailSubject;
        string emailBody;
        string[] attachments;

        string mailToShow;

        List<string> list_LUU = new List<string>();

        Dictionary<string, List<string>> dictSchoolLUU = new Dictionary<string, List<string>>();

        Encoding danish = Encoding.GetEncoding(1252);

        string forceSchoolSelect = "*Vælg skole*";



        public Form4(string _dir_summaries, string _file_LUU, string _file_summaries, string _emailFrom, DateTime _emailReceived, string _emailSubject, string _emailBody, string[] _attachments)
        {
            InitializeComponent();

            // Set constructor variables to global
            dir_summaries = _dir_summaries;
            file_LUU = _file_LUU;
            file_summaries = _file_summaries;
            emailFrom = _emailFrom;
            emailReceived = _emailReceived.ToString("dd-MM-yyyy, HH:mm:ss", CultureInfo.InvariantCulture);
            emailReceivedShortVersion = _emailReceived.ToString("yyyy-MM-dd ", CultureInfo.InvariantCulture);
            emailSubject = _emailSubject;
            emailBody = _emailBody;
            attachments = _attachments;

            // Show mail content
            richTextBox1.Text = mailToShow;

            // Set mail info to allow the user to determine the context of the summeries
            mailToShow =
                "Modtaget: " + emailReceived + Environment.NewLine +
                "Fra: " + emailFrom + Environment.NewLine +
                "Emne: " + emailSubject + Environment.NewLine + Environment.NewLine +
                emailBody;

            richTextBox1.Text = mailToShow;

            // Get LUUs to list
            list_LUU = File.ReadAllLines(file_LUU, danish).ToList();
            

            // Extract schools and LUUs from _luus
            foreach (string line in list_LUU)
            {
                string[] arrLine = line.Split(';');
                string luu = arrLine[0].Trim();
                string school = arrLine[1].Trim();

                // Add to dictionary
                if (dictSchoolLUU.ContainsKey(school))
                {
                    dictSchoolLUU[school].Add(luu);
                }
                else
                {
                    dictSchoolLUU.Add(school, new List<string>{luu});
                }
            }
            dictSchoolLUU.Add(forceSchoolSelect, null); // Force user to select a school

            // School dropdown items
            List<string> listSchoolLUU = dictSchoolLUU.Keys.ToList();
            listSchoolLUU.Sort();
            comboBox1.DataSource = listSchoolLUU;
            comboBox1.SelectedIndex = comboBox1.FindStringExact(forceSchoolSelect);

            // Load files to the datagridview
            dataGridView1.Rows.Clear();
            for (int i=0; i<attachments.Length; i++)
            {
                DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();

                row.Cells[0].Value = Path.GetFileName(attachments[i]);
                row.Cells[1].Value = attachments[i];
                row.Cells[2].Value = true;

                dataGridView1.Rows.Add(row);
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            // Skip this mail
            this.DialogResult = DialogResult.Ignore;
            this.Close();
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                List<string> schoolsArrSorted = dictSchoolLUU[comboBox1.GetItemText(comboBox1.SelectedItem)].ToList();
                schoolsArrSorted.Sort();

                comboBox2.DataSource = schoolsArrSorted;
            }
            catch
            {
                // comboBox1 (schools) will be null if the forceSchoolSelect is selected (which it is by default)
                comboBox2.DataSource = null;
            }
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
            {
                try
                {
                    System.Diagnostics.Process.Start(senderGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()); // Opening pdf with path equal to the clicked cells value as string
                }
                catch
                {
                    //
                }
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            // Check that a school has been selected in the combobox
            if (comboBox1.GetItemText(comboBox1.SelectedItem) != forceSchoolSelect)
            {
                // Get selected school and LUU
                string school = comboBox1.GetItemText(comboBox1.SelectedItem);
                string luu = comboBox2.GetItemText(comboBox2.SelectedItem);

                string luuId = get_LUUid(luu, school);

                // Write to summaries file and move the file(s) to their correct folder
                for (int i=0; i<attachments.Length; i++)
                {
                    // Check whether the file is set to be omitted in the datagridview by user
                    bool include = true;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        string dataGridViewFileName = "";
                        if (row.Cells["Column1"].Value != null)
                        {
                            dataGridViewFileName = row.Cells["Column1"].Value.ToString();
                        }
                        
                        if (dataGridViewFileName == Path.GetFileName(attachments[i]))
                        {
                            DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells["Column3"];

                            bool chkChecked = true;
                            try
                            {
                                chkChecked = Convert.ToBoolean(chk.Value);
                            }
                            catch
                            {
                                chkChecked = false;
                            }

                            if (chkChecked == true)
                            {
                                include = true;
                            }
                            else
                            {
                                include = false;
                            }
                        }
                    }

                    if (include)
                    {
                        // Place file
                        string strSummaryDirectory = dir_summaries + school + @"\" + luu + @"\";
                        Directory.CreateDirectory(strSummaryDirectory);

                        // [CL:6]
                        string attachmentFileName = Path.GetFileName(attachments[i]);
                        string strSummaryPath = strSummaryDirectory + emailReceivedShortVersion + attachmentFileName;
                        //string strSummaryPath = strSummaryDirectory + DateTime.Now.ToString("yyyyMMdd-HHmmssfff", CultureInfo.InvariantCulture) + ".pdf";

                        File.Copy(attachments[i], strSummaryPath);

                        // Shorten summary path
                        string strSummaryShortPath = strSummaryPath.Replace(Directory.GetCurrentDirectory(), "");

                        // Write to summary index
                        string receivedDate = DateTime.ParseExact(emailReceived, "dd-MM-yyyy, HH:mm:ss", CultureInfo.InvariantCulture).ToString("dd-MM-yyyy", CultureInfo.InvariantCulture);
                        string indexDate = DateTime.Now.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture);
                        string line = string.Format("{0};{1};{2};{3};{4};0{5}", luuId, receivedDate, indexDate, emailFrom, strSummaryShortPath, Environment.NewLine);
                        File.AppendAllText(file_summaries, line, danish);

                        // Delay to ensure unique file name
                        System.Threading.Thread.Sleep(100);
                    }
                }

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            // Abort this an all the following mails
            this.DialogResult = DialogResult.Abort;
            this.Close();
        }

        public string get_LUUid(string luu, string school)
        {
            string[] allLUU = File.ReadAllLines(file_LUU, danish);
            foreach (string line in allLUU)
            {
                if (line.Contains(luu) && line.Contains(school))
                {
                    string[] LUUInfo = line.Split(';');
                    return LUUInfo[3];
                }
            }

            return null;
        }
    }
}
