using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
namespace Excel
{
    public partial class Form1 : Form
    {
        SqlConnection con;
        SqlCommand cmd;
        SqlDataReader dr;
        ConnectionDB db = new ConnectionDB();
        public Form1()
        {
            InitializeComponent();

            con = new SqlConnection(db.GetConnection());
        }

        //public string conString = "Data Source=laptop-4aosrhae\\sqlexpress;Initial Catalog=DbConnection;Integrated Security=True";
        public void Loadrecords()
        {
            con.Open();
            cmd = new SqlCommand("SELECT * tb1Employees", con);
            con.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application x1App;
            Microsoft.Office.Interop.Excel.Workbook x1Workbook;
            Microsoft.Office.Interop.Excel.Worksheet x1Worksheet;
            Microsoft.Office.Interop.Excel.Range x1Range;

            int x1Row;
            string strFileName;

            openFD.Filter = "Excel Office |*.xls; *xlsx";
            openFD.ShowDialog();
            strFileName = openFD.FileName;

            if(strFileName != "")
            {
                x1App = new Microsoft.Office.Interop.Excel.Application();
                x1Workbook = x1App.Workbooks.Open(strFileName);
                x1Worksheet = x1Workbook.Worksheets["Sheet1"];
                x1Range = x1Worksheet.UsedRange;

                int i = 0;

                for (x1Row = 2; x1Row <= x1Range.Rows.Count; x1Row++)
                {
                    if (x1Range.Cells[x1Row,1].Text != "")
                    {
                        i++;
                        dgvEmployees.Rows.Add(i, x1Range.Cells[x1Row, 1].Text, x1Range.Cells[x1Row, 2].Text, x1Range.Cells[x1Row, 3].Text, x1Range.Cells[x1Row, 4].Text, x1Range.Cells[x1Row, 5].Text, x1Range.Cells[x1Row, 6].Text);
                    }
                    
                }
                x1Workbook.Close();
                
                x1App.Quit();
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Main Form1Info = new Main();
            Form1Info.Show();
        }
    }
}
