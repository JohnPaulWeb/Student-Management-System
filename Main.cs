using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public partial class Main : Form
    {

        SqlConnection sqlconn = new SqlConnection(@"Data Source=LAPTOP-4AOSRHAE\SQLEXPRESS; Initial Catalog=TestDatabase; Integrated Security=True");
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            sqlconn.Open();
            SqlCommand sqlcmd = sqlconn.CreateCommand();
            sqlcmd.CommandType = CommandType.Text;
            sqlcmd.CommandText = "select * from [TestTable]";
            sqlcmd.ExecuteNonQuery();
            DataTable datatable = new DataTable();
            SqlDataAdapter sqladapter = new SqlDataAdapter(sqlcmd);
            sqladapter.Fill(datatable);
            dataGridView1.DataSource = datatable;
            sqlconn.Close();
        }
    }
}
