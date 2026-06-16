using System;
using System.Drawing;
using System.Windows.Forms;
using System.Management;
using System.Collections;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter08
{
    public partial class Recipe08_18 : Form
    {
        public Recipe08_18()
        {
            InitializeComponent();
        }

        private void cmdRefresh_Click(object sender, EventArgs e)
        {
            // Select all the outstanding print jobs.
            string query = "SELECT * FROM Win32_PrintJob";
            using (ManagementObjectSearcher jobQuery =
              new ManagementObjectSearcher(query))
            {
                using (ManagementObjectCollection jobs = jobQuery.Get())
                {
                    // Add the jobs in the queue to the list box.
                    lstJobs.Items.Clear();
                    txtJobInfo.Text = "";
                    foreach (ManagementObject job in jobs)
                    {
                        lstJobs.Items.Add(job["JobID"]);
                    }
                }
            }
        }

        private void Recipe08_18_Load(object sender, EventArgs e)
        {
            cmdRefresh_Click(null, null);
        }

        // This helper method performs a WMI query and returns the
        // WMI job for the currently selected list box item.
        private ManagementObject GetSelectedJob()
        {
            try
            {
                // Select the matching print job.
                string query = "SELECT * FROM Win32_PrintJob " +
                  "WHERE JobID='" + lstJobs.Text + "'";
                ManagementObject job = null;
                using (ManagementObjectSearcher jobQuery =
                  new ManagementObjectSearcher(query))
                {
                    ManagementObjectCollection jobs = jobQuery.Get();
                    IEnumerator enumerator = jobs.GetEnumerator();
                    enumerator.MoveNext();
                    job = (ManagementObject)enumerator.Current;
                }
                return job;
            }
            catch (InvalidOperationException)
            {
                // the Current property of the enumerator is invalid
                return null;
            }
        }

        private void lstJobs_SelectedIndexChanged(object sender, EventArgs e)
        {
            ManagementObject job = GetSelectedJob();
            if (job == null)
            {
                txtJobInfo.Text = "";
                return;
            }

            // Display job information.
            StringBuilder jobInfo = new StringBuilder();
            jobInfo.AppendFormat("Document: {0}", job["Document"].ToString());
            jobInfo.Append(Environment.NewLine);
            jobInfo.AppendFormat("DriverName: {0}", job["DriverName"].ToString());
            jobInfo.Append(Environment.NewLine);
            jobInfo.AppendFormat("Status: {0}", job["Status"].ToString());
            jobInfo.Append(Environment.NewLine);
            jobInfo.AppendFormat("Owner: {0}", job["Owner"].ToString());
            jobInfo.Append(Environment.NewLine);

            jobInfo.AppendFormat("PagesPrinted: {0}", job["PagesPrinted"].ToString());
            jobInfo.Append(Environment.NewLine);
            jobInfo.AppendFormat("TotalPages: {0}", job["TotalPages"].ToString());

            if (job["JobStatus"] != null)
            {
                txtJobInfo.Text += Environment.NewLine;
                txtJobInfo.Text += "JobStatus: " + job["JobStatus"].ToString();
            }
            if (job["StartTime"] != null)
            {
                jobInfo.Append(Environment.NewLine);
                jobInfo.AppendFormat("StartTime: {0}", job["StartTime"].ToString());
            }

            txtJobInfo.Text = jobInfo.ToString();
        }

        private void cmdPause_Click(object sender, EventArgs e)
        {
            if (lstJobs.SelectedIndex == -1) return;
            ManagementObject job = GetSelectedJob();
            if (job == null) return;

            // Attempt to pause the job.
            int returnValue = Int32.Parse(
              job.InvokeMethod("Pause", null).ToString());

            // Display information about the return value.
            if (returnValue == 0)
            {
                MessageBox.Show("Successfully paused job.");
            }
            else
            {
                MessageBox.Show("Unrecognized return value when pausing job.");
            }
        }

        private void cmdResume_Click(object sender, EventArgs e)
        {
            if (lstJobs.SelectedIndex == -1) return;
            ManagementObject job = GetSelectedJob();
            if (job == null) return;

            if ((Int32.Parse(job["StatusMask"].ToString()) & 1) == 1)
            {
                // Attempt to resume the job.
                int returnValue = Int32.Parse(
                  job.InvokeMethod("Resume", null).ToString());

                // Display information about the return value.
                if (returnValue == 0)
                {
                    MessageBox.Show("Successfully resumed job.");
                }
                else if (returnValue == 5)
                {
                    MessageBox.Show("Access denied.");
                }
                else
                {
                    MessageBox.Show(
                      "Unrecognized return value when resuming job.");
                }
            }
        }
    }
}