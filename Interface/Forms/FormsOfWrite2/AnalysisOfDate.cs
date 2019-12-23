using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Interface.Forms.FormsOfWrite2
{
    public partial class AnalysisOfDate : Form
    {
        public AnalysisOfDate()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView2.DataSource != null)
            {
                //System.Data.DataTable sss = new System.Data.DataTable;
                //sss.Rows.Clear();
                //dataGridView1.DataSource=sss;


                System.Data.DataTable dt = (System.Data.DataTable)dataGridView2.DataSource;

                dt.Rows.Clear();

                dataGridView2.DataSource = dt;

            }

            else
            {

                dataGridView2.Rows.Clear();

            }

            foreach (KeyValuePair<string, Sensor> kvp in Global.sensors)
            {
                DataGridViewRow row = new DataGridViewRow();
                int index = dataGridView2.Rows.Add(row);
                dataGridView2.Rows[index].Cells[1].Value = kvp.Value.name;
                dataGridView2.Rows[index].Cells[2].Value = kvp.Value.serialNumber;
                dataGridView2.Rows[index].Cells[3].Value = kvp.Value.rangeMin + "-" + kvp.Value.rangeMax;
                dataGridView2.Rows[index].Cells[4].Value = kvp.Value.type;
                dataGridView2.Rows[index].Cells[5].Value = kvp.Value.uncertainty;
                dataGridView2.Rows[index].Cells[6].Value = kvp.Value.certificateNo;
                for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                {
                    dataGridView2.Rows[i].Cells[0].Value = i + 1;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            #region 删除后台word占用。
            System.Diagnostics.Process myproc = new System.Diagnostics.Process();
            //得到所有打开的进程
            try
            {
                foreach (Process thisproc in Process.GetProcessesByName("WINWORD"))
                {
                    if (!thisproc.CloseMainWindow())
                    {
                        thisproc.Kill();
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("杀死" + "WINWORD" + "失败！");
            }
            #endregion
            try
            {
                object filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName;
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document wordDoc;
                wordDoc = wordApp.Documents.Open(filename);
                wordDoc.ActiveWindow.Visible = false;//打开word

                Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[3];//检索表格

                for (int j = 0; j < dataGridView2.ColumnCount - 1; j++)
                {
                    for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                    {
                        nowtable.Cell(i + 2, j + 2).Range.InsertAfter(dataGridView2[j + 1, i].Value.ToString());//填充表格
                    }
                }
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;

                MessageBox.Show("导入成功。");
            }
            catch (Exception)
            {
                MessageBox.Show("失败。");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //OpenFileDialog fd = new OpenFileDialog();//首先根据打开文件对话框，选择excel表格
            //fd.Filter = "表格|*.xls|所有文件(*.*)|*.*";//打开文件对话框筛选器
            //string strPath;//文件完整的路径名
            //if (fd.ShowDialog() == DialogResult.OK)
            //{
            try
            {
                //strPath = fd.FileName;
                string strPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\bin\\" + Global.dataSourceName;
                string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strPath + ";extended properties='excel 12.0; HDR=NO;IMEX=1';";//关键是红色区域
                OleDbConnection Con = new OleDbConnection(strCon);//建立连接
                Con.Open();

                System.Data.DataTable sheetsName = Con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字

                string firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字

                string sql = string.Format("SELECT * FROM [{0}]", firstSheetName); //查询字符串 //string sql = string.Format("SELECT * FROM [{0}] WHERE [日期] is not null", firstSheetName); //查询字符串
                OleDbCommand Cmd = new OleDbCommand(sql, Con);//建立要执行的命令
                OleDbDataAdapter da = new OleDbDataAdapter(Cmd);//建立数据适配器
                DataSet ds = new DataSet();//新建数据集
                da.Fill(ds, "shyman");//把数据适配器中的数据读到数据集中的一个表中（此处表名为shyman，可以任取表名）
                //指定datagridview1的数据源为数据集ds的第一张表（也就是shyman表），也可以写ds.Table["shyman"]
                dataGridView1.DataSource = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);//捕捉异常
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            object filename = Application.StartupPath + "\\bin\\" + Global.templateName;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = false;//打开word
            Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[17];//检索表格
            if (comboBox1.Text == Convert.ToString(1))
            {
                nowtable.Cell(2, 1).Range.InsertAfter(Convert.ToString(dataGridView1.Rows[0].Cells[1].Value));
                nowtable.Cell(2, 2).Range.InsertAfter(textBox2.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(2))
            {
                nowtable.Cell(3, 1).Range.InsertAfter(Convert.ToString(dataGridView1.Rows[0].Cells[2].Value));
                nowtable.Cell(3, 2).Range.InsertAfter(textBox2.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(3))
            {
                nowtable.Cell(4, 1).Range.InsertAfter(Convert.ToString(dataGridView1.Rows[0].Cells[3].Value));
                nowtable.Cell(4, 2).Range.InsertAfter(textBox2.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(4))
            {
                nowtable.Cell(5, 1).Range.InsertAfter(Convert.ToString(dataGridView1.Rows[0].Cells[4].Value));
                nowtable.Cell(5, 2).Range.InsertAfter(textBox2.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(5))
            {
                nowtable.Cell(6, 1).Range.InsertAfter(Convert.ToString(dataGridView1.Rows[0].Cells[5].Value));
                nowtable.Cell(6, 2).Range.InsertAfter(textBox2.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(6))
            {
                nowtable.Cell(7, 1).Range.InsertAfter(Convert.ToString(dataGridView1.Rows[0].Cells[6].Value));
                nowtable.Cell(7, 2).Range.InsertAfter(textBox2.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(7))
            {
                nowtable.Cell(8, 1).Range.InsertAfter(Convert.ToString(dataGridView1.Rows[0].Cells[7].Value));
                nowtable.Cell(8, 2).Range.InsertAfter(textBox2.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(8))
            {
                nowtable.Cell(9, 1).Range.InsertAfter(Convert.ToString(dataGridView1.Rows[0].Cells[8].Value));
                nowtable.Cell(9, 2).Range.InsertAfter(textBox2.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(9))
            {
                nowtable.Cell(10, 1).Range.InsertAfter(Convert.ToString(dataGridView1.Rows[0].Cells[9].Value));
                nowtable.Cell(10, 2).Range.InsertAfter(textBox2.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(10))
            {
                nowtable.Cell(11, 1).Range.InsertAfter(Convert.ToString(dataGridView1.Rows[0].Cells[10].Value));
                nowtable.Cell(11, 2).Range.InsertAfter(textBox2.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            object filename = Application.StartupPath + "\\bin\\" + Global.templateName;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = false;//打开word
            Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[17];//检索表格

            for (int i = 0; i < dataGridView4.RowCount - 1; i++)
            {
                for (int j = 0; j < dataGridView4.ColumnCount; j++)
                {
                    nowtable.Cell(i + 2, j + 1).Range.InsertAfter(dataGridView4[j, i].Value.ToString());//填充表格
                }
            }

            Microsoft.Office.Interop.Word.Table nowtable1 = wordDoc.Tables[8];//检索表格
            for (int i = 0; i < 26; i++)
            {
                nowtable1.Cell(i + 3, 1).Range.InsertAfter(dataGridView1[0, i + 3].Value.ToString());
                nowtable1.Cell(i + 3, 2).Range.InsertAfter(dataGridView1[1, i + 3].Value.ToString());
                nowtable1.Cell(i + 3, 3).Range.InsertAfter(dataGridView1[0, i + 3].Value.ToString());
                nowtable1.Cell(i + 3, 4).Range.InsertAfter(dataGridView1[2, i + 3].Value.ToString());
                nowtable1.Cell(i + 3, 5).Range.InsertAfter(dataGridView1[0, i + 3].Value.ToString());
                nowtable1.Cell(i + 3, 6).Range.InsertAfter(dataGridView1[3, i + 3].Value.ToString());
            }

            Microsoft.Office.Interop.Word.Table nowtable2 = wordDoc.Tables[9];//检索表格
            for (int i = 0; i < 26; i++)
            {
                nowtable2.Cell(i + 3, 1).Range.InsertAfter(dataGridView1[0, i + 3].Value.ToString());
                nowtable2.Cell(i + 3, 2).Range.InsertAfter(dataGridView1[4, i + 3].Value.ToString());
                nowtable2.Cell(i + 3, 3).Range.InsertAfter(dataGridView1[0, i + 3].Value.ToString());
                nowtable2.Cell(i + 3, 4).Range.InsertAfter(dataGridView1[5, i + 3].Value.ToString());
                nowtable2.Cell(i + 3, 5).Range.InsertAfter(dataGridView1[0, i + 3].Value.ToString());
                nowtable2.Cell(i + 3, 6).Range.InsertAfter(dataGridView1[6, i + 3].Value.ToString());
            }

            Microsoft.Office.Interop.Word.Table nowtable3 = wordDoc.Tables[10];//检索表格
            for (int i = 0; i < 26; i++)
            {
                nowtable3.Cell(i + 3, 1).Range.InsertAfter(dataGridView1[0, i + 3].Value.ToString());
                nowtable3.Cell(i + 3, 2).Range.InsertAfter(dataGridView1[7, i + 3].Value.ToString());
                nowtable3.Cell(i + 3, 3).Range.InsertAfter(dataGridView1[0, i + 3].Value.ToString());
                nowtable3.Cell(i + 3, 4).Range.InsertAfter(dataGridView1[8, i + 3].Value.ToString());
                nowtable3.Cell(i + 3, 5).Range.InsertAfter(dataGridView1[0, i + 3].Value.ToString());
                nowtable3.Cell(i + 3, 6).Range.InsertAfter(dataGridView1[9, i + 3].Value.ToString());
            }
            wordDoc.Save();
            wordApp.Quit();
            wordApp = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
