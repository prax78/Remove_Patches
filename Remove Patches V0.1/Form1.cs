using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;

namespace Remove_Patches_V0._1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Enabled = false;
            button1.Enabled = false;
            DateTime InstalledOn=DateTime.Parse(dateTimePicker1.Value.ToString());
            String SelectedDate = $"{InstalledOn.Date.ToString("M/d/yyyy").Replace("-", "/")}";
            // MessageBox.Show($"Patches Installed on {SelectedDate}");
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("\\\\.\\root\\cimv2", $"SELECT * FROM Win32_QuickFixEngineering"))
            {
                ManagementObjectCollection Patches = searcher.Get();
                foreach (ManagementObject patch in Patches)
                {

                    //MessageBox.Show(patch["InstalledOn"].ToString());
                    if (patch["InstalledOn"].ToString() == SelectedDate)
                    {
                        dataGridView1.Rows.Add(null, patch["HotFixID"], patch["InstalledOn"], patch["Description"]);
                    }
                }
            }
            if(dataGridView1.Rows.Count>0)
            {
                button2.Enabled = true;
                button3.Enabled=true;
            }
            MessageBox.Show($"{dataGridView1.Rows.Count.ToString()} Patches for {SelectedDate}","",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            button3.Enabled = false;
            double counter = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Show();
            int totalrows = dataGridView1.RowCount;
            string message = string.Empty;
            System.Diagnostics.Stopwatch st = new System.Diagnostics.Stopwatch();
           
            double totalCheckedBox = 0;
            foreach (DataGridViewRow it in dataGridView1.Rows)
            {

                bool isSelected = Convert.ToBoolean(it.Cells["CheckBox"].Value);
                if(isSelected)
                {
                    totalCheckedBox++;
                }
            }
            st.Start();
                foreach (DataGridViewRow item in dataGridView1.Rows)
            {

                bool isSelected = Convert.ToBoolean(item.Cells["CheckBox"].Value);
                String cmd;
                if(Environment.Is64BitOperatingSystem)
                {
                    cmd = "c:\\windows\\sysnative\\wusa.exe";
                }
                else
                {
                    cmd = "c:\\windows\\system32\\wusa.exe";
                }
                if (isSelected && totalCheckedBox>0)
                {
                    String PatchToInstall = item.Cells["KB"].Value.ToString().Replace("KB", "");
                    //MessageBox.Show(PatchToInstall);
                    using (System.Diagnostics.Process pr = new System.Diagnostics.Process())
                    {
                        System.Diagnostics.ProcessStartInfo prinfo = new System.Diagnostics.ProcessStartInfo(cmd, $"/uninstall /quiet /norestart /kb:{PatchToInstall}");
                        pr.StartInfo = prinfo;
                        pr.StartInfo.UseShellExecute = true;
                        pr.StartInfo.CreateNoWindow = true;


                        await Task.Run(() => {

                            BeginInvoke(new Action(() =>
                            {
                                counter++;
                              
                                progressBar1.Value = (int)(counter /totalCheckedBox * 100);
                                label2.Text = $"Status: Uninstalling {PatchToInstall} using wusa.exe\n{counter}\\{totalCheckedBox} done\nProgress: {progressBar1.Value} %\nETA:{TimeSpan.FromSeconds((st.Elapsed.TotalSeconds/counter)*(totalCheckedBox-counter)).ToString(@"hh\:mm\:ss")} ";
                            }));
                            pr.Start();
                            pr.WaitForExit();

                        });
                        

                      
                       
                            
                        
                       


                    }

                    //message += Environment.NewLine;
                    //message += item.Cells["KB"].Value;
                }


            }
                
            label2.Text = $"Status:Completed\nTotal Time: {TimeSpan.FromSeconds( st.Elapsed.TotalSeconds).ToString(@"hh\:mm\:ss")} ";
            st.Stop();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            foreach (DataGridViewRow itm in dataGridView1.Rows)
            {
                
                itm.Cells["CheckBox"].Value = true;
            }
        }

      
    }
}
