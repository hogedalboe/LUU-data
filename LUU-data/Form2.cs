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
    public partial class Form2 : Form
    {
        /// <summary>
        /// Mode = 0: Create consultant
        /// Mode = 1: Edit consultant
        /// 
        /// </summary>

        //////////////////////////////// Globals /////////////////////////////////////////
        int mode = 0;

        List<string> local_list_LUU = new List<string>();
        string file_IUConsultants = "";
        string file_LUU = "";

        DataGridViewRow consultantRow;

        Encoding danish = Encoding.GetEncoding(1252);

        // Consultant info (mode = 1)
        string name;
        string initials;
        string email;
        string phone;

        public Form2(List<string> list_LUU, List<string> list_IUConsultants, string _file_IUConsultants, string _file_LUU, int _mode = 0, DataGridViewRow _consultantRow = null)
        {
            InitializeComponent();

            // Set Form2 variables
            local_list_LUU = list_LUU;
            file_IUConsultants = _file_IUConsultants;
            file_LUU = _file_LUU;
            mode = _mode;
            consultantRow = _consultantRow;

            // Load schools to combobox1
            List<string> filter_list_schools = new List<string>();
            for (int i = 0; i < list_LUU.Count; i++)
            {
                if (list_LUU[i] != "")
                {
                    string[] LUUInfo = list_LUU[i].Split(';');

                    if (!filter_list_schools.Contains(LUUInfo[1].Trim(';')))
                    {
                        filter_list_schools.Add(LUUInfo[1].Trim(';'));
                    }
                }
            }
            filter_list_schools.Sort();
            comboBox1.DataSource = filter_list_schools;

            // Load consultant info (if in edit mode)
            if (mode == 1)
            {
                // Set personal info
                name = consultantRow.Cells["Column1"].Value.ToString();
                initials = consultantRow.Cells["Column2"].Value.ToString();
                email = consultantRow.Cells["Column3"].Value.ToString();
                phone = consultantRow.Cells["Column4"].Value.ToString();

                // Personal info on form
                textBox1.Text = name;
                textBox2.Text = initials;
                textBox3.Text = email;
                textBox4.Text = phone;

                // LUU
                if (consultantRow.Cells["Column5"].Value != null)
                {
                    string fullConsultantLUU = consultantRow.Cells["Column5"].Value.ToString();
                    string[] semiConsultantLUU = fullConsultantLUU.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None); // https://stackoverflow.com/questions/1547476/easiest-way-to-split-a-string-on-newlines-in-net
                    foreach (string tmpConsultantLUU in semiConsultantLUU)
                    {
                        if (tmpConsultantLUU != "")
                        {
                            string[] specificConsultantLUU = tmpConsultantLUU.Split('|');

                            DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();

                            row.Cells[0].Value = specificConsultantLUU[0].Trim();
                            row.Cells[1].Value = specificConsultantLUU[1].Trim();

                            dataGridView1.Rows.Add(row);
                        }
                    }
                }
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            // Write new consultant to data files
            if (textBox1.Text != "" && // Name
                textBox3.Text != "") // Email
            {
                // Delete the current consultant from the data file so it can be overwritten
                if (mode == 1)
                {
                    // Search string for consultant to remove
                    string removeConsultant = string.Format("{0};{1};{2};{3};", name, initials, email, phone);

                    string[] consultants = File.ReadAllLines(file_IUConsultants, danish);

                    File.WriteAllText(file_IUConsultants, string.Empty, danish);

                    for (int i = 0; i < consultants.Length; i++)
                    {
                        if (!consultants[i].Contains(removeConsultant))
                        {
                            File.AppendAllText(file_IUConsultants, consultants[i] + Environment.NewLine, danish);
                        }
                    }
                }

                // Write to file with consultant info
                File.AppendAllText(file_IUConsultants,
                    textBox1.Text + ";" + // Name
                    textBox2.Text + ";" + // Initials
                    textBox3.Text + ";" + // Email
                    textBox4.Text + ";", danish); // Phone

                // Add LUU to the consultant
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells["Skole"].Value != null && row.Cells["LUU"].Value != null)
                    {
                        // Get LUU id of the current datagridview row
                        string school = row.Cells["Skole"].Value.ToString();
                        string luu = row.Cells["LUU"].Value.ToString();
                        string LUUid = getLUUId(school, luu);

                        // Write LUU id to the consultant [CL:1]
                        File.AppendAllText(file_IUConsultants, LUUid + "|", danish);

                        /*
                        //Omitted due to the need to add more than one consultant per LUU. The consultant-LUU relation is no longer maintained in the file_LUU [CL:1]
                        // Write consultant to file with LUU info (check if LUU already has another consultant attached)
                        string[] arrLUU = File.ReadAllLines(file_LUU, danish);
                        for (int i = 0; i < arrLUU.Length; i++)
                        {
                            if (arrLUU[i].Contains(LUUid))
                            {
                                string[] LUUinfo = arrLUU[i].Split(';');

                                // Check if another consultant is already attached to the LUU
                                if (LUUinfo[2] != textBox1.Text && LUUinfo[2] != "") //
                                {
                                DialogResult dialogResult = MessageBox.Show(LUUinfo[2] + " er allerede tilknyttet '" + luu + "' på " + school + "." + Environment.NewLine + Environment.NewLine + "Vil du erstatte " + LUUinfo[2] + " med " + textBox1.Text + "?", "Udskift konsulent", MessageBoxButtons.YesNo);
                                if (dialogResult == DialogResult.Yes)
                                {
                                    // Remove the LUU id from former consultant(s)
                                    string text = File.ReadAllText(file_IUConsultants, danish);
                                    text = text.Replace(LUUid + "|", "");
                                    File.WriteAllText(file_IUConsultants, text, danish);

                                    // Add the LUU id to the consultant's profile if no other consultant is attached to the LUU
                                    File.AppendAllText(file_IUConsultants, LUUid + "|", danish);

                                    // Replace former consultant's name with the new consultant's name in the LUU
                                    arrLUU[i] = LUUinfo[0] + ";" + LUUinfo[1] + ";" + textBox1.Text + ";" + LUUid;
                                }
                                else if (dialogResult == DialogResult.No)
                                {
                                    //
                                }
                                }
                                else
                                {
                                // Add the LUU id to the consultant's profile if no other consultant is attached to the LUU
                                File.AppendAllText(file_IUConsultants, LUUid + "|", danish);

                                // Add the consultants name to the LUU
                                arrLUU[i] = LUUinfo[0] + ";" + LUUinfo[1] + ";" + textBox1.Text + ";" + LUUid;
                                }
                            }
                        }
                        

                        // Write array changes to LUU file
                        File.WriteAllText(file_LUU, string.Empty, danish);
                        foreach (string LUUline in arrLUU)
                        {
                            File.AppendAllText(file_LUU, LUUline + Environment.NewLine, danish);
                        }
                        */
                    }
                }

                // Add new line to consultant file
                File.AppendAllText(file_IUConsultants, Environment.NewLine, danish);

                // Close form
                this.Close();
            }
            else
            {
                MessageBox.Show("Navn og e-mail skal være udfyldt.");
            }
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Load LUU to combobox2
            string school = comboBox1.SelectedValue.ToString();
            List<string> filter_list_LUU = new List<string>();

            for (int i = 0; i < local_list_LUU.Count; i++)
            {
                if (local_list_LUU[i] != "")
                {
                    string[] LUUInfo = local_list_LUU[i].Split(';');

                    if (!filter_list_LUU.Contains(LUUInfo[0].Trim(';')))
                    {
                        bool includeLUU = true;

                        // Check if the LUU is under the selected school (combobox1)
                        if (school != null && LUUInfo[1].Trim(';') != school)
                        {
                            includeLUU = false;
                        }

                        if (includeLUU)
                        {
                            filter_list_LUU.Add(LUUInfo[0].Trim(';'));
                        }
                    }
                }
            }
            filter_list_LUU.Sort();
            comboBox2.DataSource = filter_list_LUU;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            string school = comboBox1.SelectedValue.ToString();
            string luu = comboBox2.SelectedValue.ToString();

            // Check if LUU already exists on the consultant
            bool LUUExists = false;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["Skole"].Value != null && row.Cells["LUU"].Value != null)
                {
                    if (row.Cells["Skole"].Value.ToString() == school && row.Cells["LUU"].Value.ToString() == luu)
                    {
                        LUUExists = true;
                    }
                }
            }

            if (!LUUExists)
            {
                DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();

                row.Cells[0].Value = school;
                row.Cells[1].Value = luu;

                dataGridView1.Rows.Add(row);
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Rows.RemoveAt(dataGridView1.CurrentRow.Index);
            }
            catch
            {
                //
            }
        }

        public string getLUUId(string school, string luu)
        {
            for (int i = 0; i < local_list_LUU.Count; i++)
            {
                if (local_list_LUU[i] != "")
                {
                    if (local_list_LUU[i].Contains(school) && local_list_LUU[i].Contains(luu))
                    {
                        // Return LUU id
                        string[] LUUInfo = local_list_LUU[i].Split(';');
                        return LUUInfo[3];
                    }
                }
            }

            return null;
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
