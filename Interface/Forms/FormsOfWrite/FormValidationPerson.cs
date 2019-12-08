using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Interface
{
    public partial class FormValidationPerson : Form
    {
        public FormValidationPerson()
        {
            InitializeComponent();
            for (int i = 0; i < 4; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                int index = dataGridView1.Rows.Add(row);
                dataGridView1.Rows[index].Cells[0].Value = "";
                dataGridView1.Rows[index].Cells[1].Value = "";
            }
            for (int i = 0; i < 3; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                int index = dataGridView2.Rows.Add(row);
                dataGridView2.Rows[index].Cells[0].Value = "";
                dataGridView2.Rows[index].Cells[1].Value = "";
            }
            {
                DataTable dt = new DataTable();
                DataColumn dc = new DataColumn("职责", typeof(System.String));
                dt.Columns.Add(dc);
                dt.Rows.Add("协助组织实施");
                dt.Rows.Add("组织协调、监督实施，复核关键数据");
                dt.Rows.Add("验证过程的复核");
                dt.Rows.Add("方案、报告审核");
                dataGridView3.DataSource = dt;
            }
            {
                DataTable dt2 = new DataTable();
                DataColumn dc3 = new DataColumn("验证职责", typeof(System.String));
                dt2.Columns.Add(dc3);
                dt2.Rows.Add("负责起草方案、书写报告");
                dt2.Rows.Add("协助实施验证、对相关人员进行培训");
                dt2.Rows.Add("协助验证项目整体管理、协调及内审");
                dataGridView4.DataSource = dt2;
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();

        }

        private void FormValidationPerson_Load(object sender, EventArgs e)
        {
            object filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = false;//打开word

            Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[1];//检索表格

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    nowtable.Cell(i + 2, j + 1).Range.InsertAfter(dataGridView1[j, i].Value.ToString());//填充表格
                }
            }
            wordDoc.Close();
            wordApp.Quit();
            wordApp = null;


        }



        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            #region 删除后台word占用。
            //
            Process myProcess = new Process();
            Process[] wordProcess = Process.GetProcessesByName("winword");
            foreach (Process pro in wordProcess) //这里是找到那些没有界面的Word进程
            {
                IntPtr ip = pro.MainWindowHandle;

                string str = pro.MainWindowTitle; //发现程序中打开跟用户自己打开的区别就在这个属性
                //用户打开的str 是文件的名称，程序中打开的就是空字符串
                if (str != "冷库验证项目模拟报告模板")
                {
                    pro.Kill();
                }
            }
            #endregion
            try
            {
                object filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName;
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document wordDoc;
                wordDoc = wordApp.Documents.Open(filename);
                wordDoc.ActiveWindow.Visible = false;//打开word

                Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[1];//检索表格

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        nowtable.Cell(i + 2, j + 1).Range.InsertAfter(dataGridView1[j, i].Value.ToString());//填充表格
                    }
                }

                Microsoft.Office.Interop.Word.Table nowtable2 = wordDoc.Tables[2];//检索表格

                for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                {
                    for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    {
                        nowtable2.Cell(i + 2, j + 1).Range.InsertAfter(dataGridView2[j, i].Value.ToString());//填充表格
                    }
                }


                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;

                MessageBox.Show("导入成功。");
                this.Close();

            }
            catch (Exception)
            {
                MessageBox.Show("导入失败。");
            }
        }


    }
}
