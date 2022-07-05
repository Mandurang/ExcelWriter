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
        private readonly Repo repo = new Repo();

        public Form1()
        {
            InitializeComponent();
            //dgv.DataSource = repo.RepoIntal();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = "dfdfre";
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

            string[] data = { "abx", "asd", "qwe" ,"fdf"};

            string fileName = "ExcelData";

            LoadData(fileName, columns, data);
        }

        public void LoadData(string fileName ,string[] columns, string[] data)
        {
            addGridParam(columns, dgv);
            addGridParam(data, dgv);

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

        public void addGridParam(string[] N, DataGridView Grid)
             
        {

            //пока столбцов не будет достаточное количество добавляем их

            while (N.Length > Grid.ColumnCount)

            {

                //если колонок нехватает добавляем их пока их будет хватать

                Grid.Columns.Add("", "");

            }

            //заполняем строку

            Grid.Rows.Add(N);

        }

        //CreateExcelAsync(string fileName, string[] columns, string[] data)



    }
}
