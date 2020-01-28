using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LUU_data
{
    public partial class Form3 : Form
    {
        //////////////////////////////// Globals /////////////////////////////////////////

        Encoding danish = Encoding.GetEncoding(1252);

        string dir_summaries;

        string file_IUConsultants;
        string file_LUU;

        int mode = 0;

        string currentId;
        string currentLuu;
        string currentSchool;
        string currentConsultant;

        /// <summary>
        /// mode=0: Add LUU
        /// mode=1: Edit LUU
        /// </summary>


        public Form3(string _dir_summaries, List<string> list_LUU, List<string> list_IUConsultants, string _file_IUConsultants, string _file_LUU, int _mode=0, string _LUUid=null, string _LUU=null, string _school=null, string _consultant=null)
        {
            InitializeComponent();

            // Set Form3 variables
            dir_summaries = _dir_summaries;
            file_IUConsultants = _file_IUConsultants;
            file_LUU = _file_LUU;
            mode = _mode;
            currentId = _LUUid;
            currentLuu = _LUU;
            currentSchool = _school;
            currentConsultant = _consultant;

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

            // Load consultants to combobox2
            List<string> filter_list_consultants = new List<string>();
            for (int i = 0; i < list_IUConsultants.Count; i++)
            {
                if (list_IUConsultants[i] != "")
                {
                    string[] consultantInfo = list_IUConsultants[i].Split(';');

                    if (!filter_list_consultants.Contains(consultantInfo[0].Trim(';')))
                    {
                        filter_list_consultants.Add(consultantInfo[0].Trim(';'));
                    }
                }
            }
            filter_list_consultants.Sort();
            filter_list_consultants.Add(""); // Empty selection
            comboBox2.DataSource = filter_list_consultants;

            // Add mode
            if (mode == 0)
            {
                button3.Hide();
            }

            // Edit mode
            if (mode == 1)
            {
                textBox1.Text = currentLuu;
                comboBox1.SelectedIndex = comboBox1.FindStringExact(currentSchool);
                comboBox2.SelectedIndex = comboBox2.FindStringExact(currentConsultant);
            }
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                // Disable adding of new school
                checkBox2.Checked = false;
                textBox2.Enabled = false;
            }
            else
            {
                // Enable adding of new school
                checkBox2.Checked = true;
                textBox2.Enabled = true;
            }
        }

        private void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                // Disable choosing existing school
                checkBox1.Checked = false;
                comboBox1.Enabled = false;
            }
            else
            {
                // Enable choosing existing school
                checkBox1.Checked = true;
                comboBox1.Enabled = true;
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            // A name must be inputted
            if (textBox1.Text != "")
            {
                // Either existing or new school must be selected/inputted
                if (comboBox1.SelectedIndex > -1 || textBox2.Text != "")
                {
                    string[] luu = File.ReadAllLines(file_LUU, danish);

                    string luuName = textBox1.Text;

                    string schoolName = "";
                    if (textBox2.Text != "")
                    {
                        schoolName = textBox2.Text;
                    }
                    else
                    {
                        schoolName = comboBox1.GetItemText(comboBox1.SelectedItem);
                    }

                    // Check if the LUU already exists on the school
                    bool luuExists = false;
                    for (int i=0; i<luu.Length; i++)
                    {
                        string[] luuInfo = luu[i].Split(';');

                        if (schoolName == luuInfo[1] && luuName == luuInfo[0] && currentId != luuInfo[3])
                        {
                            luuExists = true;
                        }
                    }

                    // If the LUU already exists on the school, don't do anything but prompting the user for a change of LUU name
                    if (!luuExists)
                    {
                        if (mode == 1)
                        {
                            // Delete the existing LUU in file, so that it can be overwritten
                            File.WriteAllText(file_LUU, string.Empty, danish);

                            foreach (string line in luu)
                            {
                                if (!line.Contains(currentId))
                                {
                                    File.AppendAllText(file_LUU, line + Environment.NewLine, danish);
                                }
                            }

                            // Remove LUU from consultant file
                            string completeConsultant = File.ReadAllText(file_IUConsultants, danish);
                            completeConsultant = completeConsultant.Replace(currentId + "|", "");
                            File.WriteAllText(file_IUConsultants, completeConsultant, danish);
                        }

                        string consultantName = comboBox2.GetItemText(comboBox2.SelectedItem);

                        // Add LUU with new id if it being created for the first time
                        string id = "";
                        if (mode == 0)
                        {
                            string date = DateTime.Now.ToString("HHmmssddMMyyyy");

                            string user = Environment.UserName;

                            id = string.Format("{0}-{1}-{2}-{3}", Clean(luuName), Clean(schoolName), date, user);
                        }
                        // Otherwise use the same id
                        else if (mode == 1)
                        {
                            id = currentId;
                        }

                        // Add LUU to the LUU file
                        string lineLUU = string.Format("{0};{1};{2};{3}" + Environment.NewLine, luuName, schoolName, "", id); // consultantName removed [CL:1]
                        File.AppendAllText(file_LUU, lineLUU, danish);

                        // Add LUU id to the specified consultant
                        if (consultantName != "")
                        {
                            string[] consultants = File.ReadAllLines(file_IUConsultants, danish);

                            for (int i = 0; i < consultants.Length; i++)
                            {
                                if (consultants[i].Contains(consultantName))
                                {
                                    consultants[i] = consultants[i] + id + "|";
                                }
                            }

                            File.WriteAllText(file_IUConsultants, string.Empty);
                            foreach (string line in consultants)
                            {
                                File.AppendAllText(file_IUConsultants, line + Environment.NewLine, danish);
                            }
                        }

                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("LUUet eksisterer allerede på skolen.");
                    }
                }
                else
                {
                    MessageBox.Show("Kun feltet med IU-konsulent må være blank");
                }
            }
            else
            {
                MessageBox.Show("Kun feltet med IU-konsulent må være blank");
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public string Clean(string str)
        {
            string s = Regex.Replace(str, "[^ÆØÅæøåA-Za-z]", "");

            s = s.ToLower();

            return s;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Er du sikker på, at du vil slette dette LUU?", "Slet LUU", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                // Delete LUU in LUU file
                string[] arrLUU = File.ReadAllLines(file_LUU, danish);
                List<string> listLUU = new List<string>(arrLUU);

                for (int i = listLUU.Count-1; i >= 0; i--)
                {
                    if (listLUU[i].Contains(currentId))
                    {
                        listLUU.RemoveAt(i);
                    }
                }

                File.WriteAllText(file_LUU, string.Empty, danish);
                foreach (string line in listLUU)
                {
                    File.AppendAllText(file_LUU, line + Environment.NewLine, danish);
                }

                // Remove LUU from consultant file
                string completeConsultant = File.ReadAllText(file_IUConsultants, danish);
                completeConsultant = completeConsultant.Replace(currentId + "|", "");
                File.WriteAllText(file_IUConsultants, completeConsultant, danish);

                this.Close();
            }
        }
    }
}
