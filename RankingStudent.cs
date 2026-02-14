using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public partial class RankingStudent : Form
    {
        private string mfilename = null;
        public RankingStudent()
        {
            InitializeComponent();
        }

        private void btnRanking_Click(object sender, EventArgs e)
        {
            double[] R = new double[14];

            R[0] = Convert.ToDouble(txtNet.Text);
            R[1] = Convert.ToDouble(txtJava.Text);
            R[2] = Convert.ToDouble(txtNet.Text);
            R[3] = Convert.ToDouble(txtNet.Text);
            R[4] = Convert.ToDouble(txtNet.Text); 
            R[5] = Convert.ToDouble(txtNet.Text);
            R[6] = Convert.ToDouble(txtNet.Text);
            R[7] = Convert.ToDouble(txtNet.Text);

            R[8] = (R[0] + R[1] + R[2] + R[3] + R[4] + R[5] + R[6] + R[7]) / 8;
            R[9] = R[0] + R[1] + R[2] + R[3] + R[4] + R[5] + R[6] + R[7];

            string average = Convert.ToString(R[8]);
            string total = Convert.ToString(R[9]);

            lblAverage.Text = average;
            lblScore.Text = total;
            if (R[9] >= 700)
            {
                lblRanking.Text = "1st";
                txtProgression.Text = "Student Completed";
            }
            else if (R[9] >= 600)
            {
                lblRanking.Text = "2nd";
                txtProgression.Text = "Student Not Completed";
            }
            else if (R[9] >= 500)
            {
                lblRanking.Text = "3rd";
                txtProgression.Text = "Student Not Completed";
            }
            else if (R[9] >= 400)
            {
                lblRanking.Text = "4rd";
                txtProgression.Text = "Student Not Completed";
            }
            else if (R[9] >= 300)
            {
                lblRanking.Text = "5th";
                txtProgression.Text = "Student Not Completed";
            }
            else if (R[9] >= 200)
            {
                lblRanking.Text = "6th";
                txtProgression.Text = "Student Not Completed";
            }
            else if (R[9] >= 100)
            {
                lblRanking.Text = "7th";
                txtProgression.Text = "Student Not Completed";
            }


        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(txtStudent_ID.Text, txtFirstname.Text, txtSurname.Text, comboBox1.Text,
                txtNet.Text, txtJava.Text, txtComputing.Text, txtPhilosophy.Text, txtPR.Text, txtBusiness.Text,
                txtBiology.Text, txtAccounting.Text, lblScore.Text, lblAverage.Text, lblRanking.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook | *.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                       //txtFileName.Text = OpenFileDialog.;
                    }
                }
            }
        }

        private void btnTranscript_Click(object sender, EventArgs e)
        {
            rtTranscript.AppendText("Student_ID \t\t\t" + txtStudent_ID.Text + "\nName: \t" + txtFirstname.Text + "\t" + txtSurname.Text + "\n");
            rtTranscript.AppendText("Net :\t\t\t\t" + (txtNet.Text) + "\n\n");
            rtTranscript.AppendText("Java :\t\t\t\t" + (txtJava.Text) + "\n");
            rtTranscript.AppendText("Computing :\t\t\t" + (txtComputing.Text) + "\n");
            rtTranscript.AppendText("Philosophy :\t\t\t" + (txtPhilosophy.Text) + "\n");
            rtTranscript.AppendText("PR  :\t\t\t" + (txtPR.Text) + "\n");
            rtTranscript.AppendText("Business :\t\t\t" + (txtBusiness.Text) + "\n");
            rtTranscript.AppendText("Biology :\t\t\t" + (txtBiology.Text) + "\n");
            rtTranscript.AppendText("Accounting :\t\t\t" + (txtAccounting.Text) + "\n");
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            rtTranscript.Clear();
            dataGridView1.Rows.Clear();
            txtStudent_ID.Clear();
            txtNet.Clear();
            txtJava.Clear();
            txtComputing.Clear();
            txtPhilosophy.Clear();
            txtPR.Clear();
            txtBusiness.Clear();
            txtBiology.Clear();
            txtAccounting.Clear();
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            rtTranscript.Clear();
            dataGridView1.Rows.Clear();
            txtStudent_ID.Clear();
            txtNet.Clear();
            txtJava.Clear();
            txtComputing.Clear();
            txtPhilosophy.Clear();
            txtPR.Clear();
            txtBusiness.Clear();
            txtBiology.Clear();
            txtAccounting.Clear();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            this.dataGridView1.EditMode = dataGridView1.EditMode;
        }
    }
}
