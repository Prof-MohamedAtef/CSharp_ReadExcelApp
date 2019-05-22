using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Reflection;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel; // Add Reference : COM Tab : Microsoft Excel Object Library ...


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataSet1.Clear();
            if (textBox1.Text==string.Empty)
            {
                MessageBox.Show("Insert Sheet Name In the Text Box");   
            }
            else
            {
                string strsheet = textBox1.Text;
                OpenFileDialog openfiledialouge1 = new OpenFileDialog();
                openfiledialouge1.Filter = "Spread Sheets(*.*)|*.xlsx|All Files (*.*)|*.*";

                if (openfiledialouge1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        string strfilename = openfiledialouge1.FileName;
                        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strfilename + ";Extended Properties=Excel 12.0");
                        OleDbDataAdapter adapter = new OleDbDataAdapter("select * from [TimeTable$]", con);//["+strsheet.ToString()+"$]", con);
                        adapter.Fill(dataSet1);
                        dataGridView1.DataSource = dataSet1.Tables[0];
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        private void ExcelWBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataSet1.Clear();
            string strsheet = textBox1.Text;
            OpenFileDialog openfiledialouge1 = new OpenFileDialog();
            openfiledialouge1.Filter = "Spread Sheets(*.*)|*.xlsx|All Files (*.*)|*.*";

            if (openfiledialouge1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    string strfilename = openfiledialouge1.FileName;
                    OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strfilename + ";Extended Properties=Excel 12.0");
                    OleDbDataAdapter adapter = new OleDbDataAdapter("select * from [TimeTable$]", con);//["+strsheet.ToString()+"$]", con);
                    adapter.Fill(dataSet1);
                    dataGridView1.DataSource = dataSet1.Tables[0];
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
    }
}
