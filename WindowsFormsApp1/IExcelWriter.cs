using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public interface IExcelWriter
    {
        void LoadData(string fileName, string[] columns, string[] data, DataGridView Grid);
        void addGridColumns(string[] stringData, DataGridView Grid);
        void addGridData(string[] data, DataGridView grid); 
    }
}
