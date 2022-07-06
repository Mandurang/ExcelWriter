using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private readonly IExcelWriter _writer;

        public Form1(IExcelWriter writer)
        {
            InitializeComponent();
            _writer = writer;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = "New file";
            Excel.Application exApp = new Excel.Application();

            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            
            int i, j;
            for (i = 0; i <= dgv.RowCount - 1; i++)
            {
                for (j = 0; j <= dgv.ColumnCount - 1; j++)
                {
                    
                    if (dgv[j, i].Value != null)
                    wsh.Cells[i + 1, j + 1] = dgv[j, i].Value.ToString();
                }
            }
            exApp.Visible = true;
            exApp.GetSaveAsFilename(fileName);

        }

        private void Load_Click(object sender, EventArgs e)
        {
            string[] columns = { "cibergod", "is", "good", "special", "Site" };

            string[] data = { "abx", "asd", "qwe" ,"fdf", "df", "df", "dfd"};

            string fileName = "ExcelData";

            _writer.LoadData(fileName, columns, data, dgv);
        }
    }
}
