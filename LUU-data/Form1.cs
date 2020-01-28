using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Outlook.Application;

/* CHANGELOG:
 *      - 2019-10-14 [CL:1]: Added the ability to add more than one consultant per LUU. This required comprehensive changes to read/write logics on several forms.
 *      - 2019-10-15 [CL:2]: Added the ability to see which date interval was previously read from Outlook.
 *      - 2019-10-17 [CL:3]: Added the ability to delete summaries.
 *      - 2019-10-18 [CL:4]: Allowed reading of outlook to be done for only one day (before minimum two separate days were required by mistake).
 *      - 2019-10-18 [CL:5]: The ability to send updates to all relevant consultants for a summary was not included in [CL:1], so this ability was added.
 *      - 2019-11-05 [CL:6]: Summaries are now saved with their original file name and Outlook receiving date instead of just the load date (Form4).
 *      - 2019-11-05 [CL:7]: Shading every second row in each dataGridView (Form1).
 *      - 2019-11-05: All combo boxes are now loaded with alphabetically-sorted data sources.
 *      - 2019-11-13: Made some adjustments to the [CL:6] changes. Sorted comboboxes on Form4. Allowed for searching in all comboboxes across forms (AutoCompleteMode.SuggestAppend; AutoCompleteSource.ListItems;).
 *      - 2020-01-28: readCategory string set according to request by SFA ("Referater indlæst af LULU"). Also previous categories are deleted.
 *      - 2020-01-28 [CL:8]: Date interval selection correction of endDate.
 *      
 *      
 * 
 */

namespace LUU_data
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////// GLOBALS /////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////////////////////////

        // ENCODINGS
        Encoding danish = Encoding.GetEncoding(1252);

        // FILES & DIRECTORIES
        string file_latestOutlookReading = Directory.GetCurrentDirectory() + @"\data\settings\latest-outlook-reading.txt";
        string file_outlookAccount = Directory.GetCurrentDirectory() + @"\data\settings\outlook-account.txt";
        string file_outlookCategory = Directory.GetCurrentDirectory() + @"\data\settings\outlook-category.txt";
        string file_outlookSubject = Directory.GetCurrentDirectory() + @"\data\settings\outlook-subject.txt";
        string file_outlookBody = Directory.GetCurrentDirectory() + @"\data\settings\outlook-body.txt";
        string file_IUConsultants = Directory.GetCurrentDirectory() + @"\data\iu-konsulenter.txt";
        string file_LUU = Directory.GetCurrentDirectory() + @"\data\luu.txt";
        string file_summaries = Directory.GetCurrentDirectory() + @"\data\referater.txt";
        string dir_summaries = Directory.GetCurrentDirectory() + @"\data\referater\";
        string dir_temporary = Directory.GetCurrentDirectory() + @"\data\temporary\";
        string dir_backup = Directory.GetCurrentDirectory() + @"\data\backup\";
        string dir_settings = Directory.GetCurrentDirectory() + @"\data\settings\";

        // SETTINGS
        string outlookAccount;
        string outlookCategory;
        string outlookSubject;
        string outlookBody;
        string comboboxDefaultValue = " * Alle*";

        // DATA
        List<string> list_IUConsultants = new List<string>();
        List<string> list_LUU = new List<string>();
        List<string> list_summaries = new List<string>();

        // BACKUP
        string dir_currentBackup;

        private void Form1_Load(object sender, EventArgs e)
        {
            ////////////////////////////////// LAYOUT ////////////////////////////////////////
            this.MinimumSize = this.Size;
            this.MaximumSize = this.Size;

            progressBar1.Hide();

            //////////////////////////////// LOAD DATA ///////////////////////////////////////

            // Load Outlook settings;
            outlookAccount = File.ReadAllText(file_outlookAccount, danish);
            outlookCategory = File.ReadAllText(file_outlookCategory, danish);
            outlookSubject = File.ReadAllText(file_outlookSubject, danish);
            outlookBody = File.ReadAllText(file_outlookBody, danish);

            // Set minimum date for datetimepicker1
            list_summaries = File.ReadAllLines(file_summaries, danish).ToList();
            DateTime minDate = DateTime.Now;
            for (int i = 0; i < list_summaries.Count; i++)
            {
                string[] summaryInfo = list_summaries[i].Split(';');
                string strSummaryDate = summaryInfo[1].Trim(';');

                DateTime summaryDate = DateTime.ParseExact(strSummaryDate, "dd-MM-yyyy", CultureInfo.InvariantCulture);
                if (summaryDate < minDate)
                {
                    minDate = summaryDate;
                }
            }
            dateTimePicker1.Value = minDate;

            ////////////////////////////// LOAD FORM ELEMENTS ////////////////////////////////

            // Outlook settings
            textBox2.Text = outlookCategory;
            textBox1.Text = outlookSubject;
            textBox3.Text = outlookAccount;
            richTextBox1.Text = outlookBody;

            // IU Consultants in datagridview1
            load_IUConsultants();

            // LUU in datagridview2
            load_LUU();

            // Summaries in datagridview3
            load_Summaries();

            // Filter schools for datagridview3
            load_FilterSchools(true);

            // Filter LUU for datagridview3
            load_FilterLUU(true);

            //////////////////// PLACE BELOW CODE AT THE BOTTOM OF Form1_Load //////////////////

            // Make a backup of the settings directory and table files
            dir_currentBackup = dir_backup + "backup " + DateTime.Now.ToString("yyyy-MM-dd HHmmss fff", CultureInfo.InvariantCulture) + @"\";
            backup();
        }


        //////////////////////////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////// FUNCTIONS: READING //////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////////////////////////

        // This function returns a list containing the school and LUU name by the input of a LUU identifyer string.
        public string[] get_LUU(string LUUid)
        {
            string[] allLUU = File.ReadAllLines(file_LUU, danish);

            foreach (string LUU in allLUU)
            {
                if (LUU.Contains(LUUid) && LUU != "")
                {
                    string[] LUUInfo = LUU.Split(';');
                    return LUUInfo;
                }
            }

            return null;
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

        // Load IU consultants to datagridview1
        public void load_IUConsultants()
        {
            // Load IU Consultants
            list_IUConsultants = File.ReadAllLines(file_IUConsultants, danish).ToList();

            dataGridView1.Rows.Clear();
            for (int i = 0; i < list_IUConsultants.Count; i++)
            {
                if (list_IUConsultants[i] != "")
                {
                    string[] consultantInfo = list_IUConsultants[i].Split(';');

                    DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();

                    row.Cells[0].Value = consultantInfo[0].Trim(';');
                    row.Cells[1].Value = consultantInfo[1].Trim(';');
                    row.Cells[2].Value = consultantInfo[2].Trim(';');
                    row.Cells[3].Value = consultantInfo[3].Trim(';');

                    // Get the LUUs that the consultant is responsible for
                    if (consultantInfo[4] != "")
                    {
                        string[] LUUids = consultantInfo[4].Split('|');
                        string strCellLUU = "";
                        foreach (string LUUid in LUUids)
                        {
                            if (LUUid != "")
                            {
                                string[] LUUInfo = get_LUU(LUUid.Trim(';')); // Get the school and LUU name from an ID
                                strCellLUU = strCellLUU + LUUInfo[1].Trim(';') + " | " + LUUInfo[0].Trim(';') + Environment.NewLine;
                            }
                        }

                        if (strCellLUU != "")
                        {
                            row.Cells[4].Value = strCellLUU;
                        }
                    }

                    //[CL:7]
                    if (i % 2 != 0)
                    {
                        row.DefaultCellStyle.BackColor = Color.LightBlue;
                    }

                    dataGridView1.Rows.Add(row);
                }
            }
        }

        // Get consultants on LUU
        public List<string> getConsultantsOnLUU(string LUUId)
        {
            List<string> list_matchingIUConsultants = new List<string>();

            list_IUConsultants = File.ReadAllLines(file_IUConsultants, danish).ToList();

            for (int i = 0; i < list_IUConsultants.Count; i++)
            {
                if (list_IUConsultants[i] != "")
                {
                    string[] consultantInfo = list_IUConsultants[i].Split(';');

                    if (consultantInfo[4].Contains(LUUId))
                    {
                        list_matchingIUConsultants.Add(consultantInfo[0]);
                    }
                }
            }

            return list_matchingIUConsultants;
        }

        // Load LUU to datagridview2
        public void load_LUU()
        {
            // Load LUU
            list_LUU = File.ReadAllLines(file_LUU, danish).ToList();

            dataGridView2.Rows.Clear();
            for (int i = 0; i < list_LUU.Count; i++)
            {
                if (list_LUU[i] != "")
                {
                    string[] LUUInfo = list_LUU[i].Split(';');

                    DataGridViewRow row = (DataGridViewRow)dataGridView2.Rows[0].Clone();

                    row.Cells[0].Value = LUUInfo[0].Trim(';');
                    row.Cells[1].Value = LUUInfo[1].Trim(';');

                    // Get names of consultants on LUU
                    //row.Cells[2].Value = LUUInfo[2].Trim(';'); // Omitted due to adding of the ability to connect several consultants to one LUU [CL:1]
                    List<string> list_matchingIUConsultants = getConsultantsOnLUU(LUUInfo[3]);
                    string consultants = "";
                    foreach (string consultant in list_matchingIUConsultants)
                    {
                        consultants = consultants + consultant + Environment.NewLine;
                    }
                    row.Cells[2].Value = consultants;

                    //[CL:7]
                    if (i % 2 != 0)
                    {
                        row.DefaultCellStyle.BackColor = Color.LightBlue;
                    }

                    dataGridView2.Rows.Add(row);
                }
            }
        }

        // Load summaries to datagridview3
        public void load_Summaries(string filter_school = null,
            string filter_LUU = null,
            DateTime? filter_fromdate = null,
            DateTime? filter_todate = null)
        {
            // Load Summaries
            list_summaries = File.ReadAllLines(file_summaries, danish).ToList();

            dataGridView3.Rows.Clear();
            for (int i = 0; i < list_summaries.Count; i++)
            {
                if (list_summaries[i] != "")
                {
                    string[] summaryInfo = list_summaries[i].Split(';');

                    DataGridViewRow row = (DataGridViewRow)dataGridView3.Rows[0].Clone();

                    // Get school and LUU name from ID
                    string[] LUUInfo = get_LUU(summaryInfo[0].Trim(';'));
                    row.Cells[0].Value = LUUInfo[0].Trim(';'); // LUU name
                    row.Cells[1].Value = LUUInfo[1].Trim(';'); // School name

                    row.Cells[2].Value = summaryInfo[1].Trim(';'); // Summary date
                    row.Cells[3].Value = summaryInfo[2].Trim(';'); // Import date
                    row.Cells[4].Value = summaryInfo[3].Trim(';'); // Sender

                    row.Cells[5].Value = Directory.GetCurrentDirectory() + @summaryInfo[4].Trim(';'); // Summary path button

                    // Advised (y/n)
                    if (summaryInfo[5].Trim(';') == "1")
                    {
                        row.Cells[6].Value = true;
                    }
                    else
                    {
                        row.Cells[6].Value = false;
                    }

                    // Filter bool to set
                    bool addRow = true;

                    // Summary date to DateTime object
                    DateTime summaryDate = DateTime.ParseExact(summaryInfo[1].Trim(';'), "dd-MM-yyyy", CultureInfo.InvariantCulture);

                    // Filter by school
                    if (filter_school != null && filter_school != comboboxDefaultValue)
                    {
                        if (LUUInfo[1].Trim(';') != filter_school)
                        {
                            addRow = false;
                        }
                    }

                    // Filter by LUU
                    if (filter_LUU != null && filter_LUU != comboboxDefaultValue)
                    {
                        if (LUUInfo[0].Trim(';') != filter_LUU)
                        {
                            addRow = false;
                        }
                    }

                    // Filter by fromdate
                    if (filter_fromdate != null)
                    {
                        if (summaryDate < filter_fromdate)
                        {
                            addRow = false;
                        }
                    }

                    // Filter by todate
                    if (filter_todate != null)
                    {
                        if (summaryDate > filter_todate)
                        {
                            addRow = false;
                        }
                    }

                    if (addRow)
                    {
                        //[CL:7]
                        if (i % 2 != 0)
                        {
                            row.DefaultCellStyle.BackColor = Color.LightBlue;
                        }

                        dataGridView3.Rows.Add(row);
                    }
                }
            }
        }

        // Filter datagridview3 (summaries)
        public void filter_Summaries()
        {
            string school = null;
            if (comboBox1.SelectedValue != null)
            {
                school = comboBox1.SelectedValue.ToString();
            }

            string LUU = null;
            if (comboBox2.SelectedValue != null)
            {
                LUU = comboBox2.SelectedValue.ToString();
            }

            DateTime from = dateTimePicker1.Value;
            DateTime to = dateTimePicker2.Value;

            load_Summaries(school, LUU, from, to);
        }

        // Load schools to dropdown (filter for datagridview3)
        public void load_FilterSchools(bool selectLastIndex=false)
        {
            List<string> filter_list_schools = new List<string>();

            for (int i=0; i<list_LUU.Count; i++)
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

            // Add empty selection item
            filter_list_schools.Sort();
            filter_list_schools.Add(comboboxDefaultValue);

            comboBox1.DataSource = filter_list_schools;

            // Select last item in combobox
            if (selectLastIndex)
            {
                comboBox1.SelectedIndex = comboBox1.Items.Count - 1;
            }
        }

        // Load LUUs to dropdown (filter for datagridview3)
        public void load_FilterLUU(bool selectLastIndex = false, string school=null)
        {
            List<string> filter_list_LUU = new List<string>();

            for (int i = 0; i < list_LUU.Count; i++)
            {
                if (list_LUU[i] != "")
                {
                    string[] LUUInfo = list_LUU[i].Split(';');

                    if (!filter_list_LUU.Contains(LUUInfo[0].Trim(';')))
                    {
                        bool includeLUU = true;

                        // Check if the LUU is under the selected school (combobox1)
                        if (school != null && LUUInfo[1].Trim(';') != school && school != comboboxDefaultValue)
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

            // Add empty selection item
            filter_list_LUU.Sort();
            filter_list_LUU.Add(comboboxDefaultValue);

            comboBox2.DataSource = filter_list_LUU;

            // Select last item in combobox
            if (selectLastIndex)
            {
                comboBox2.SelectedIndex = comboBox2.Items.Count - 1;
            }
        }

        private string GetSenderSMTPAddress(MailItem mail)
        {
            // Credit: https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-smtp-address-of-the-sender-of-a-mail-item

            string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            if (mail == null)
            {
                throw new ArgumentNullException();
            }
            if (mail.SenderEmailType == "EX")
            {
                AddressEntry sender =
                    mail.Sender;
                if (sender != null)
                {
                    //Now we have an AddressEntry representing the Sender
                    if (sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry || 
                        sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                    {
                        //Use the ExchangeUser object PrimarySMTPAddress
                        ExchangeUser exchUser = sender.GetExchangeUser();
                        if (exchUser != null)
                        {
                            return exchUser.PrimarySmtpAddress;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    else
                    {
                        return sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                    }
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return mail.SenderEmailAddress;
            }
        }

        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Credit: https://docs.microsoft.com/en-us/dotnet/standard/io/how-to-copy-directories

            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, true);
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

        public void backup()
        {
            Directory.CreateDirectory(dir_currentBackup);

            File.Copy(file_IUConsultants, dir_currentBackup + Path.GetFileName(file_IUConsultants));
            File.Copy(file_LUU, dir_currentBackup + Path.GetFileName(file_LUU));
            File.Copy(file_summaries, dir_currentBackup + Path.GetFileName(file_summaries));

            DirectoryCopy(dir_settings, dir_currentBackup + @"\settings\", true);

            // Remove the oldest backup
            if (Directory.GetDirectories(dir_backup).Length > 5)
            {
                FileSystemInfo fileInfo = new DirectoryInfo(dir_backup).GetFileSystemInfos().OrderBy(fi => fi.CreationTime).First(); // Credit: https://stackoverflow.com/questions/44690815/how-to-delete-oldest-folder-created-from-local-disk-using-c-sharp
                Directory.Delete(fileInfo.FullName, true);
            }
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////// FUNCTIONS: WRITING //////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////////////////////////

        public void alignConsultantsWithLUU()
        {
            /*
             * Omitted 14-10-2019 due to the need to add more than one consultant per LUU. Therefore, no alignment is needed because only the file_IUConsultant contains the relation [CL:1]

            string[] consultants = File.ReadAllLines(file_IUConsultants, danish);
            string[] luu = File.ReadAllLines(file_LUU, danish);

            // Iterate the LUUs to check whether they don't have a consultant attached any more
            for (int j = 0; j < luu.Length; j++)
            {
                bool haveConsultant = false;

                string[] luuInfo = luu[j].Split(';');

                // Current LUU id
                string luuId = luuInfo[3];

                for (int i = 0; i < consultants.Length; i++)
                {
                    if (consultants[i].Contains(luuId))
                    {
                        haveConsultant = true;
                    }
                }

                if (!haveConsultant)
                {                    
                    // Remove the attached consultant
                    try
                    {
                        luu[j] = luu[j].Replace(luuInfo[2], "");
                    }
                    catch
                    {

                    }

                    // DEBUG
                    //MessageBox.Show(string.Format("{0} was removed from '{1}'", luuInfo[2], luuInfo[3]));
                }
            }

            // Write array changes to LUU file
            File.WriteAllText(file_LUU, string.Empty, danish);
            foreach (string LUUline in luu)
            {
                File.AppendAllText(file_LUU, LUUline + Environment.NewLine, danish);
            }
            */
        }


        //////////////////////////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////// EVENT HANDLERS //////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////////////////////////

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                var senderGrid = (DataGridView)sender;

                // Handle clicking on button that should open the file represented on the line
                if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
                {
                    try
                    {
                        Process.Start(senderGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()); // Opening pdf with path equal to the clicked cells value as string
                    }
                    catch
                    {
                        //
                    }
                }

                // Handle toggling of status checkbox
                if (senderGrid.Columns[e.ColumnIndex] is DataGridViewCheckBoxColumn && e.RowIndex >= 0)
                {
                    DataGridViewCheckBoxCell cell = senderGrid.Rows[e.RowIndex].Cells[e.ColumnIndex] as DataGridViewCheckBoxCell;

                    string summaryPath = senderGrid.Rows[e.RowIndex].Cells[5].Value.ToString().Replace(Directory.GetCurrentDirectory(), "");

                    if (Convert.ToBoolean(cell.Value) == true)
                    {
                        // Unchecked, therefore the summary should be sent to the relevant consultant
                        cell.Value = false;
                        changeStatus(summaryPath, Convert.ToBoolean(cell.Value));
                    }
                    else
                    {
                        // Checked, don't send anything
                        cell.Value = true;
                        changeStatus(summaryPath, Convert.ToBoolean(cell.Value));
                    }

                    void changeStatus(string subStringOnLineToChange, bool status)
                    {
                        string[] summaries = File.ReadAllLines(file_summaries, danish);

                        for (int i = 0; i < summaries.Length; i++)
                        {
                            if (summaries[i].Contains(summaryPath))
                            {
                                string[] summary = summaries[i].Split(';');

                                string nLine = "";

                                for (int j = 0; j < summary.Length; j++)
                                {
                                    if (j == 0)
                                    {
                                        nLine = summary[j] + ";";
                                    }
                                    else if (j != 5)
                                    {
                                        nLine = nLine + summary[j] + ";";
                                    }
                                    else
                                    {
                                        if (status)
                                        {
                                            nLine = nLine + "1";
                                        }
                                        else
                                        {
                                            nLine = nLine + "0";
                                        }
                                    }
                                }

                                // Write changes to summaries file
                                string FullSummaries = File.ReadAllText(file_summaries, danish);
                                FullSummaries = FullSummaries.Replace(summaries[i], nLine);
                                File.WriteAllText(file_summaries, FullSummaries, danish);
                            }
                        }
                    }
                }
            }
            catch
            {

            }
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            filter_Summaries();

            // Only show LUU under the selected school
            load_FilterLUU(true, comboBox1.SelectedValue.ToString());
        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            filter_Summaries();
        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            filter_Summaries();
        }

        private void DateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            filter_Summaries();
        }

        private void DateTimePicker1_MouseDown(object sender, MouseEventArgs e)
        {
            var senderDtp = (DateTimePicker)sender;

            senderDtp.Value = DateTime.Now;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Form2 addConsultant = new Form2(list_LUU, list_IUConsultants, file_IUConsultants, file_LUU, 0);
            addConsultant.ShowDialog();

            // Update the LUU file
            //alignConsultantsWithLUU(); // [CL:1]

            // Reload datagridviews
            load_IUConsultants();
            load_LUU();
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridViewRow currentRow = dataGridView1.CurrentRow;

                Form2 editConsultant = new Form2(list_LUU, list_IUConsultants, file_IUConsultants, file_LUU, 1, currentRow);
                editConsultant.ShowDialog();

                // Update the LUU file
                //alignConsultantsWithLUU(); // [CL:1]

                // Reload datagridviews
                load_IUConsultants();
                load_LUU();
            }
            catch
            {

            }
        }

        private void DataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Er du sikker på, at du vil slette " + dataGridView1.CurrentRow.Cells["Column1"].Value.ToString() + "?", "Slet konsulent", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    // Delete consultant in file_consultant
                    string name = dataGridView1.CurrentRow.Cells["Column1"].Value.ToString();
                    string initials = dataGridView1.CurrentRow.Cells["Column2"].Value.ToString();
                    string email = dataGridView1.CurrentRow.Cells["Column3"].Value.ToString();
                    string phone = dataGridView1.CurrentRow.Cells["Column4"].Value.ToString();
                    string removeConsultant = string.Format("{0};{1};{2};{3};", name, initials, email, phone); // Search string for consultant to remove

                    string[] consultants = File.ReadAllLines(file_IUConsultants, danish);

                    File.WriteAllText(file_IUConsultants, string.Empty, danish);

                    for (int i = 0; i < consultants.Length; i++)
                    {
                        if (!consultants[i].Contains(removeConsultant))
                        {
                            File.AppendAllText(file_IUConsultants, consultants[i] + Environment.NewLine, danish);
                        }
                    }

                    // Update the LUU file
                    //alignConsultantsWithLUU(); // [CL:1]

                    // Reload datagridviews
                    load_IUConsultants();
                    load_LUU();
                }
                else if (dialogResult == DialogResult.No)
                {
                    //
                }
            }
            catch
            {
                //
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Form3 addLUU = new Form3(dir_summaries, list_LUU, list_IUConsultants, file_IUConsultants, file_LUU);
            addLUU.ShowDialog();

            // Reload datagridviews
            load_IUConsultants();
            load_LUU();
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridViewRow currentRow = dataGridView2.CurrentRow;
                string luu = currentRow.Cells["dataGridViewTextBoxColumn1"].Value.ToString();
                string school = currentRow.Cells["dataGridViewTextBoxColumn2"].Value.ToString();
                string consultant = currentRow.Cells["dataGridViewTextBoxColumn3"].Value.ToString();

                string currentId = get_LUUid(luu, school);

                Form3 editLUU = new Form3(dir_summaries, list_LUU, list_IUConsultants, file_IUConsultants, file_LUU, 1, currentId, luu, school, consultant);
                editLUU.ShowDialog();

                // Reload datagridviews
                load_IUConsultants();
                load_LUU();
            }
            catch
            {

            }
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            Process.Start(dir_summaries);
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            // Have the user delimit date interval to read mail items
            Form5 chooseDateInterval = new Form5();
            var dResult = chooseDateInterval.ShowDialog();

            // Handle success
            if (dResult == DialogResult.OK)
            {
                // Date interval for mail items
                DateTime startDate = chooseDateInterval.startDate;
                DateTime endDate = chooseDateInterval.endDate.AddHours(23).AddMinutes(59).AddSeconds(59); // [CL:8]

                // Outlook App
                Application outlookApplication = null;
                NameSpace outlookNamespace = null;
                MAPIFolder inboxFolder = null;
                Items mailItems = null;

                try
                {
                    // Handle abort
                    bool abort = false;

                    // Handle skip
                    List<string> avoidMailDublicates = new List<string>();

                    outlookApplication = new Application();
                    outlookNamespace = outlookApplication.GetNamespace("MAPI");
                    /*inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);*/

                    // Set progress bar maximum
                    int max = 0;
                    foreach (Store accountStore in outlookNamespace.Stores)
                    {
                        try
                        {
                            inboxFolder = accountStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                        }
                        catch
                        {
                            //
                        }
                        if (inboxFolder != null)
                        {
                            max = max + inboxFolder.Items.Count;
                        }
                    }
                    progressBar1.Show();
                    progressBar1.Maximum = max;
                    progressBar1.Value = 0;

                    // Iterate the inbox of each Outlook account
                    foreach (Store accountStore in outlookNamespace.Stores)
                    {
                        try
                        {
                            // Get current account's inbox
                            inboxFolder = accountStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                        }
                        catch
                        {

                        }

                        if (inboxFolder != null)
                        {
                            mailItems = inboxFolder.Items;

                            foreach (object mailObj in mailItems)
                            {
                                if (!abort)
                                {
                                    if (mailObj is MailItem && mailObj != null)
                                    {
                                        MailItem item = (MailItem)mailObj;

                                        // Only handle mails within selected date interval (Form5)
                                        if (item.ReceivedTime >= startDate && item.ReceivedTime <= endDate)
                                        {
                                            string itemCategories = item.Categories;

                                            if (itemCategories != null && itemCategories.Contains(outlookCategory))
                                            {
                                                // Add a hardcoded category to indicate whether an e-mail has been read before
                                                string readCategory = "Referater indlæst af LULU";

                                                // Ignore previously read items
                                                if (!itemCategories.Contains(readCategory))
                                                {
                                                    // Get attachments
                                                    List<string> attachments = new List<string>();
                                                    var attachmentsRaw = item.Attachments;
                                                    string[] extensionsArray = { ".pdf", ".doc", ".docx" };
                                                    if (attachmentsRaw.Count > 0)
                                                    {
                                                        for (int i = 1; i <= attachmentsRaw.Count; i++)
                                                        {
                                                            if (extensionsArray.Any(attachmentsRaw[i].FileName.Contains))
                                                            {
                                                                // Save attachment in temporary directory
                                                                string tmpPath = dir_temporary + attachmentsRaw[i].FileName;
                                                                string tmpFileName = Path.GetFileNameWithoutExtension(tmpPath);
                                                                string tmpUnique = DateTime.Now.ToString("HHmmssfff", CultureInfo.InvariantCulture);
                                                                tmpPath = tmpPath.Replace(tmpFileName, tmpFileName + " (" + tmpUnique + ")");

                                                                System.Threading.Thread.Sleep(200);
                                                                attachmentsRaw[i].SaveAsFile(tmpPath);

                                                                // If the file is a Word document, convert it to a pdf file
                                                                string tmpPathExtension = Path.GetExtension(tmpPath);
                                                                if (tmpPathExtension == ".docx" || tmpPathExtension == ".doc")
                                                                {
                                                                    // Word app
                                                                    Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                                                                    app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                                                                    app.Visible = false;

                                                                    Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(tmpPath);

                                                                    // Save as pdf
                                                                    tmpPath = tmpPath.Replace(tmpPathExtension, ".pdf");
                                                                    try
                                                                    {
                                                                        doc.SaveAs2(tmpPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                                                                    }
                                                                    catch
                                                                    {
                                                                        //
                                                                    }

                                                                    // Cleanup Word
                                                                    doc.Close(false);
                                                                    app.Quit(false, false, false);
                                                                    Marshal.ReleaseComObject(app);
                                                                }

                                                                // Attachments is later passed to Form4
                                                                attachments.Add(tmpPath);
                                                            }
                                                        }
                                                    }

                                                    // Only relevant to handle input if there are recognized summary attachments
                                                    if (attachments.Count > 0)
                                                    {
                                                        // Set a unique identifier to avoid showing the same mail several times
                                                        string uniqueIdentifier =
                                                            "Modtaget: " + item.ReceivedTime.ToString("dd-MM-yyyy, HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine +
                                                            "Fra: " + item.SenderEmailAddress + Environment.NewLine +
                                                            "Emne: " + item.Subject + Environment.NewLine + Environment.NewLine +
                                                            item.Body;

                                                        // Make sure the sender's address is SMTP and not EX
                                                        string from = GetSenderSMTPAddress(item);
                                                        if (from == null)
                                                        {
                                                            from = item.SenderEmailAddress;
                                                        }

                                                        if (!avoidMailDublicates.Contains(uniqueIdentifier))
                                                        {
                                                            avoidMailDublicates.Add(uniqueIdentifier);

                                                            // Open Form4 to have the user determine what to do with the attachments
                                                            Form4 inputSummary = new Form4(dir_summaries, file_LUU, file_summaries, from, item.ReceivedTime, item.Subject, item.Body, attachments.ToArray());
                                                            var result = inputSummary.ShowDialog();

                                                            // Handle success
                                                            if (result == DialogResult.OK)
                                                            {
                                                                // Indicate that the e-mail has been read
                                                                item.Categories = readCategory;
                                                            }
                                                            // Handle skip
                                                            else if (result == DialogResult.Ignore)
                                                            {
                                                                //
                                                            }
                                                            // Handle abort
                                                            else if (result == DialogResult.Abort)
                                                            {
                                                                abort = true;
                                                            }
                                                        }
                                                    }

                                                    Marshal.ReleaseComObject(item);
                                                }
                                            }
                                        }
                                    }
                                }
                                // Progress bar
                                try
                                {
                                    progressBar1.Value++;
                                }
                                catch
                                {
                                    //MessageBox.Show(progressBar1.Value.ToString() + " / " + progressBar1.Maximum.ToString());
                                }
                            }
                        }
                    }

                    // Register successful load of Outlook to show in next period delimitation [CL:2]
                    File.WriteAllText(file_latestOutlookReading, String.Empty, danish);
                    File.AppendAllText(file_latestOutlookReading, startDate.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture) + Environment.NewLine, danish);
                    File.AppendAllText(file_latestOutlookReading, endDate.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture) + Environment.NewLine, danish);
                    File.AppendAllText(file_latestOutlookReading, DateTime.Today.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture) + Environment.NewLine, danish);
                    File.AppendAllText(file_latestOutlookReading, Environment.UserName, danish);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    // Cleanup
                    ReleaseComObject(inboxFolder);
                    ReleaseComObject(mailItems);
                    ReleaseComObject(outlookNamespace);
                    ReleaseComObject(outlookApplication);

                    load_Summaries();

                    progressBar1.Hide();

                    // Clean up the temporary directory
                    DirectoryInfo di = new DirectoryInfo(dir_temporary);
                    foreach (FileInfo file in di.GetFiles())
                    {
                        try
                        {
                            file.Delete();
                        }
                        catch
                        {

                        }
                    }
                }

                void ReleaseComObject(object obj)
                {
                    if (obj != null)
                    {
                        Marshal.ReleaseComObject(obj);
                        obj = null;
                    }
                }
            }
            // Handle abort
            else if (dResult == DialogResult.Abort)
            {
                // Don't do anything
            }
        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(file_outlookCategory, textBox2.Text, danish);
            outlookCategory = File.ReadAllText(file_outlookCategory, danish);
        }

        Account GetAccountForFolder(MAPIFolder folder, Microsoft.Office.Interop.Outlook.Application app)
        {
            // https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-account-for-a-folder

            // Obtain the store on which the folder resides.
            Store store = folder.Store;

            // Enumerate the accounts defined for the session.
            foreach (Account account in app.Session.Accounts)
            {
                // Match the DefaultStore.StoreID of the account
                // with the Store.StoreID for the currect folder.
                if (account.DeliveryStore.StoreID == store.StoreID)
                {
                    // Return the account whose default delivery store
                    // matches the store of the given folder.
                    return account;
                }
            }
            // No account matches, so return null.
            return null;
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(file_outlookSubject, textBox1.Text, danish);
            outlookSubject = File.ReadAllText(file_outlookSubject, danish);
        }

        private void RichTextBox1_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(file_outlookBody, richTextBox1.Text, danish);
            outlookBody = File.ReadAllText(file_outlookBody, danish);
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            load_Summaries();
            load_IUConsultants();

            bool notifiedAboutAccount = false;

            // Send relevant summaries to each consultant
            for (int i=0; i<list_IUConsultants.Count; i++)
            {
                string[] lineConsultant = list_IUConsultants[i].Split(';');

                string consultantEmailAddress = lineConsultant[2];

                string infoToSend = "<ul>"; // HTML unordered list

                for (int j = 0; j < list_summaries.Count; j++)
                {
                    string[] lineSummary = list_summaries[j].Split(';');

                    // Check whether the summary should be announced
                    if (lineSummary[5] == "0")
                    {
                        // Check if the summary belongs to the consultant
                        if (list_IUConsultants[i].Contains(lineSummary[0]))
                        {
                            // Shared directory path to summary file
                            string summaryPath = Directory.GetCurrentDirectory() + lineSummary[4];

                            // LUU name
                            string[] LUUInfo = get_LUU(lineSummary[0]);
                            string summaryName = LUUInfo[0] + ", " + LUUInfo[1] + ", " + lineSummary[1];

                            infoToSend = string.Format("{0}<li><a href='{1}'>{2}<a/></li>", infoToSend, summaryPath, summaryName);

                            // Check whether another consultant should also be notified by iterating THE REST of the consultants [CL:5]
                            bool noOtherConsultant = true;
                            for (int g=i+1; g<list_IUConsultants.Count; g++)
                            {
                                // Is the current summary present on the line of the sub-iterated consultants
                                if (list_IUConsultants[g].Contains(lineSummary[0]))
                                {
                                    noOtherConsultant = false;
                                }
                            }

                            // Indicate that the summary has been announced to the consultant
                            if (noOtherConsultant)
                            {
                                string newLine = "";
                                for (int k = 0; k < lineSummary.Length; k++)
                                {
                                    if (k == 0)
                                    {
                                        newLine = lineSummary[k] + ";";
                                    }
                                    else if (k != 5)
                                    {
                                        newLine = newLine + lineSummary[k] + ";";
                                    }
                                    else
                                    {
                                        newLine = newLine + "1";
                                    }
                                }
                                list_summaries[j] = newLine;
                            }
                        }
                    }
                }

                infoToSend = infoToSend + "</ul>";

                // Send mail
                if (infoToSend != "<ul></ul>")
                {
                    string tmpOutlookBody = outlookBody.Replace("{0}", infoToSend);

                    // Outlook application
                    Application oApp = new Application();
                    MailItem oMsg = (MailItem)oApp.CreateItem(OlItemType.olMailItem);
                    oMsg.BodyFormat = OlBodyFormat.olFormatHTML;

                    // Send from a specific account
                    bool outlookAccountFound = false;
                    try
                    {
                        foreach (Account account in oApp.Session.Accounts)
                        {
                            if (account.SmtpAddress == outlookAccount)
                            {
                                oMsg.SendUsingAccount = account;
                                outlookAccountFound = true;
                            }
                        }
                    }
                    catch
                    {
                        //
                    }
                    finally
                    {
                        if (!outlookAccountFound && !notifiedAboutAccount)
                        {
                            notifiedAboutAccount = true;
                            MessageBox.Show("Det ser ikke ud til, at du kan sende fra den ønskede mail-adresse '" + outlookAccount + "." + Environment.NewLine + "Din primære mail-adresse anvendes i stedet for.");
                        }
                    }

                    // Create mail
                    oMsg.Subject = outlookSubject;
                    oMsg.Recipients.Add(consultantEmailAddress);
                    System.Threading.Thread.Sleep(2000);
                    oMsg.Display();
                    oMsg.HTMLBody = tmpOutlookBody;
                }
            }

            // Write the updated list of summaries to the summary file (0 changed to 1)
            File.WriteAllText(file_summaries, string.Empty, danish);
            File.WriteAllLines(file_summaries, list_summaries, danish);

            load_Summaries();
        }

        private void TextBox3_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(file_outlookAccount, textBox3.Text, danish);
            outlookAccount = textBox3.Text;
        }

        private void DataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Button9_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Er du sikker på, at du vil (af)markere alle referater?" + 
                Environment.NewLine + "Handlingen kan ikke fortrydes.", "(Af)marker alle referater", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {
                string[] summaries = File.ReadAllLines(file_summaries, danish);

                bool allSummariesSent = true;
                bool noSummariesSent = true;

                // Find out which way to toggle the checkboxes
                for (int i = 0; i < summaries.Length; i++)
                {
                    string[] lineSummary = summaries[i].Split(';');

                    if (lineSummary[5] == "0")
                    {
                        allSummariesSent = false;
                    }
                    else if (lineSummary[5] == "1")
                    {
                        noSummariesSent = false;
                    }
                }

                // Mixed values: Check all checkboxes
                if (!allSummariesSent && !allSummariesSent)
                {
                    toggleSummaries(true);
                }
                // All unchecked: Check all checkboxes
                else if (noSummariesSent)
                {
                    toggleSummaries(true);
                }
                // All checked: Uncheck all checkboxes
                else if (allSummariesSent)
                {
                    toggleSummaries(false);
                }

                void toggleSummaries(bool check)
                {
                    for (int i = 0; i < summaries.Length; i++)
                    {
                        string[] lineSummary = summaries[i].Split(';');

                        string newLine = "";

                        if (check)
                        {
                            lineSummary[5] = "1";
                        }
                        else
                        {
                            lineSummary[5] = "0";
                        }

                        for (int j = 0; j < lineSummary.Length; j++)
                        {
                            if (j < lineSummary.Length - 1)
                            {
                                newLine = newLine + lineSummary[j] + ";";
                            }
                            else
                            {
                                newLine = newLine + lineSummary[j];
                            }
                        }

                        summaries[i] = newLine;
                    }

                    // Write to summaries file and load summaries to Form1
                    File.WriteAllLines(file_summaries, summaries, danish);
                    load_Summaries();
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                //
            }
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            // [CL:3]

            try
            {
                DialogResult dialogResult = MessageBox.Show("Er du sikker på, at du vil slette dette referat?" + Environment.NewLine + "Handlingen kan ikke fortrydes.", "Slet referat?", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.Yes)
                {
                    DataGridViewRow currentRow = dataGridView3.CurrentRow;
                    string luu = currentRow.Cells["dataGridViewTextBoxColumn4"].Value.ToString();
                    string school = currentRow.Cells["dataGridViewTextBoxColumn5"].Value.ToString();
                    string summaryPath = currentRow.Cells["dataGridViewLinkColumn1"].Value.ToString();

                    // LUU id
                    string luuId = get_LUUid(luu, school);

                    // Ommit (delete) summary
                    string[] summaries = File.ReadAllLines(file_summaries, danish);

                    File.WriteAllText(file_summaries, String.Empty, danish);

                    foreach (string summary in summaries)
                    {
                        string[] summaryColumns = summary.Split(';');

                        bool keep = true;

                        // If summary id and summary path are in the line, it is ommitted
                        if (summary.Contains(luuId) && summaryPath.Contains(summaryColumns[4]))
                        {
                            keep = false;
                        }

                        if (keep)
                        {
                            File.AppendAllText(file_summaries, summary + Environment.NewLine, danish);
                        }
                        else
                        {
                            // Delete the summary
                            try
                            {
                                File.Delete(summaryPath);
                            }
                            catch
                            {
                                // If the file cannot be deleted, then keep it and inform the user
                                File.AppendAllText(file_summaries, summary + Environment.NewLine, danish);
                                MessageBox.Show("Programmet kunne ikke få adgang til filen, da den bliver anvendt i en anden proces. Afslut denne proces og prøv igen.");
                            }
                        }
                    }

                    load_Summaries();
                    filter_Summaries();
                }
            }
            catch
            {

            }
        }
    }
}
