using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace Excel_App
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataTable table = new DataTable();
        private void Form1_Load(object sender, EventArgs e)
        {
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Processor Type", typeof(string));
            table.Columns.Add("Memory RAM", typeof(string));
            table.Columns.Add("Price", typeof(int));

            dataGridView1.DataSource = table;

        }
        private void btn_Import_Click(object sender, EventArgs e)
        {
            string[] lines = File.ReadAllLines(@"C:\Users\Пользователь\source\repos\Excel_App\FileExcel.txt");
            string[] values;

            for (int i = 0; i < lines.Length; i++)
            {
                values = lines[i].ToString().Split('/');
                string[] row = new string[values.Length];

                for (int j = 0; j < values.Length; j++)
                {
                    row[j] = values[j].Trim();

                }
                table.Rows.Add(row);
            }
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "AboutLaptops";
            
             for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
             {
                 worksheet.Cells[i, 1] = dataGridView1.Columns[i - 1].HeaderText;
             }

             for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
             {
                 for (int j = 0; j < dataGridView1.Columns.Count; j++)
                 {
                     worksheet.Cells[i + 6, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                 }
             }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = "output";
            saveFileDialog.DefaultExt = ".xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog.FileName, Type.Missing,Type.Missing,Type.Missing,Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            app.Quit();
        }

        private void btn_SumPrice_Click(object sender, EventArgs e)
        {
            int[] columnData = new int[dataGridView1.Rows.Count];
            columnData = (from DataGridViewRow row in dataGridView1.Rows
                          where row.Cells[4].FormattedValue.ToString() != string.Empty
                          select Convert.ToInt32(row.Cells[4].FormattedValue)).ToArray();
            label1.Text = columnData.Sum().ToString();

        }
    }
}
