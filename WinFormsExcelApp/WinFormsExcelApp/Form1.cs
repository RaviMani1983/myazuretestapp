using System;
using System.Data;
using System.Windows.Forms;

namespace WinFormsExcelApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public void Test()
        {
            try
            {
                string strFile = textBox1.Text;

                Microsoft.Office.Interop.Excel.Application m_XlApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbooks m_xlWrkbs = m_XlApp.Workbooks;

                Microsoft.Office.Interop.Excel.Workbook m_xlWrkb;
                m_xlWrkb = null;

                m_xlWrkb = m_xlWrkbs.Open(strFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing);

                Microsoft.Office.Interop.Excel.Worksheet excelSheet = m_xlWrkb.Sheets[1];
                Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                int rows = excelRange.Rows.Count;
                int cols = excelRange.Columns.Count;

                //Set DataTable Name and Columns Name
                DataTable myTable = new DataTable("MyDataTable");

                for (int i = 1; i <= cols; i++)
                {
                    myTable.Columns.Add(excelRange.Cells[1, i].Value2.ToString(), typeof(string));
                }

                for (int i = 2; i <= rows; i++)
                {
                    DataRow myNewRow = myTable.NewRow();
                    for (int c = 1; c <= cols; c++)
                    {
                        myNewRow[c - 1] = excelRange.Cells[i, c].Value2.ToString();
                    }
                    myTable.Rows.Add(myNewRow);
                }

                dataGridView1.DataSource = myTable;


                //m_xlWrkb.Save();
                m_xlWrkb.Close(true, strFile, null);
                m_XlApp.Quit();

                m_xlWrkbs = null;

                m_xlWrkb = null;

                m_XlApp = null;


                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);
            }
            catch (Exception exc)
            {

                MessageBox.Show(exc.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Test();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog.FileName;
            }
            else { textBox1.Text = ""; }
        }
    }
}
