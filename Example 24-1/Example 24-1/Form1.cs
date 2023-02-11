using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//first add reference
using Excel = Microsoft.Office.Interop.Excel;
//---------------------------------------------------------
namespace Example_24_1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //---------------------------------------------------------
        private void btnData_Click(object sender, EventArgs e)
        {
            dataGridView1.RowCount = 10;
            dataGridView1.ColumnCount = 10;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            for (int i = 1; i <= 10; i++)
            {
                for (int j = 1; j <= 10; j++)
                {
                    dataGridView1[j - 1, i - 1].Value = i * j;
                }
            }
        }
        //---------------------------------------------------------
        private void btnExport_Click(object sender, EventArgs e)
        {
            var excel = new Excel.Application();
            excel.Application.Workbooks.Add(true);
            int ColumnIndex = 0;
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                ColumnIndex++;
                excel.Cells[1, ColumnIndex] = col.HeaderText;
            }
            int rowIndex = 0;
            string val;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                rowIndex++;
                ColumnIndex = 0;
                foreach (DataGridViewColumn col in dataGridView1.Columns)
                {
                    ColumnIndex++;
                    if (row.Cells[ColumnIndex - 1].Value == null)
                        val = "";
                    else
                        val = row.Cells[ColumnIndex - 1].Value.ToString();
                    excel.Cells[rowIndex + 1, ColumnIndex] = val;
                }
            }
            excel.Visible = true;
        }
    }
}
