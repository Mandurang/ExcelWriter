using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public class ExcelWriter : IExcelWriter
    {
        public void addGridColumns(string[] stringData, DataGridView grid)
        {
            foreach (string str in stringData)
            {
                grid.Columns.Add(str, str);
            }
            grid.Rows.Add(stringData);
        }

        public void addGridData(string [] data, DataGridView grid)
        {
            var subarrays = data
                .Select((s, i) => new { Value = s, Index = i })
                .GroupBy(x => x.Index / grid.Columns.Count + 1)
                .Select(grp => grp.Select(x => x.Value).ToArray())
                .ToArray();

            foreach (var subarray in subarrays)
            {
                grid.Rows.Add(subarray);
            }
        }

        public void LoadData(string fileName, string[] columns, string[] data, DataGridView dgv)
        {
            addGridColumns(columns, dgv);
            addGridData(data, dgv);

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
    }
}
