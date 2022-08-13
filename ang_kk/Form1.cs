using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ang_kk
{
    public partial class Form1 : Form
    {

        string[] vs;
        string[] vj;
        string filename = "";

        public Form1()
        {
            InitializeComponent();
            
            
        }

        Excel.Application xlApp = new Excel.Application();
        int inc = 0, mode = 0;
        
        Random random = new Random();
        int tt = 1;
        private void guna2Button1_Click(object sender, EventArgs e)
        {
            if (mode == 1)
            {
                guna2TextBox1.Text = vj[inc].Substring(0, tt);
                tt++;
            }
            else if (mode == 2)
            {
                guna2TextBox1.Text = vs[inc].Substring(0, tt);
                tt++;
            }

        }

        private void en_kk()
        {
            inc = random.Next(1, vs.Length);
            guna2Button2.Text = vs[inc];
            guna2TextBox1.Text = "";
            tt = 0;

        }

        private void kk_en()
        {
            inc = random.Next(1, vj.Length);
            guna2Button2.Text = vj[inc];
            guna2TextBox1.Text = "";
            tt = 0;

        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {
            if (mode == 1)
            {
                if (vj[inc].ToLower().Trim() == guna2TextBox1.Text.ToLower().Trim())
                {
                    en_kk();
                    
                }
            }
            else
            {
                if (vs[inc].ToLower().Trim() == guna2TextBox1.Text.ToLower().Trim())
                {
                    kk_en();
                }
            }
        }

        private void tema1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mode = 1;
            en_kk();
            if (vs[inc].ToLower().Trim() == guna2TextBox1.Text.ToLower().Trim())
            {
                en_kk();
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog();





            filename = openFileDialog1.FileName;
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename);
            Excel._Worksheet xlWorksheet = xlWorkbook.ActiveSheet as Excel.Worksheet;
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            vs = new string[rowCount];
            vj = new string[rowCount];
            for (int i = 1; i < rowCount; i++)
            {
                vs[i] = Convert.ToString(xlRange.Cells[i, 1].Value);
                vj[i] = Convert.ToString(xlRange.Cells[i, 2].Value);
            }


        }

        private void tema1ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            mode = 2;
            kk_en();
            if (vj[inc].ToLower().Trim() == guna2TextBox1.Text.ToLower().Trim())
            {
                kk_en();
            }
        }


    }
}
