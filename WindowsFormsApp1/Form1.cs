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

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private readonly Repo repo = new Repo();

        public Form1()
        {
            InitializeComponent();
            dgv.DataSource = repo.RepoIntal();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();

            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            
            int i, j;
            for (i = 0; i <= dgv.RowCount - 1; i++)
            {
                for (j = 0; j <= dgv.ColumnCount - 1; j++)
                {
                    
                    wsh.Cells[i + 1, j + 1] = dgv[j, i].Value.ToString();
                }
            }
            exApp.Visible = true;
        }
    }
}
