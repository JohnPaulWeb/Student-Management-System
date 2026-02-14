using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtName.Text == "*")
            {
                MessageBox.Show("Enter Username !!", "Department");
            }
            else if (txtPassword.Text == "")
            {
                MessageBox.Show("Enter Password !!", "Department1");
            }
            else if (txtName.Text == "Department" && txtPassword.Text == "Department1")
            {
                this.Hide();
                Dashboard Dashboard = new Dashboard();
                Dashboard.ShowDialog();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           txtName.Clear();
            txtPassword.Clear();
        }
    }
}
