using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Description_LookupCode
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string path, sheet, pathcon = "";

        private void OpenBott_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog open = new OpenFileDialog();
            if (open.ShowDialog()== System.Windows.Forms.DialogResult.OK)
            {
                path = this.textBox1.Text = open.FileName;
                sheet = textBox2.Text;
                pathcon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;\"";
                OleDbConnection con = new OleDbConnection(pathcon);
                try
                {
                    OleDbDataAdapter myDA = new OleDbDataAdapter("select * from [" + sheet + "$]", con);
                    DataTable dt = new DataTable();
                    myDA.Fill(dt);
                    dataGridView1.DataSource = dt;
                }
                catch
                {
                    MessageBox.Show("Error Sheet Name");
                }

            }
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }
    }
}
