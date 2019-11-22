using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace Interface
{
    public partial class FormAnalysisOfData : Form
    {
        public AnyCAD.Presentation.RenderWindow3d m_RenderView ;
        //public AnyCAD.Presentation.RenderWindow3d RenderView
        //{
        //    set
        //    {
        //        m_RenderView = value;
        //    }
        //}

        public FormAnalysisOfData()
        {
            InitializeComponent();
            var container = this.panel1;

            m_RenderView = new AnyCAD.Presentation.RenderWindow3d();
            m_RenderView.Size = container.ClientSize;
            m_RenderView.Dock = System.Windows.Forms.DockStyle.Fill;
            container.Controls.Add(m_RenderView);
        }//初始化

        public static System.Data.DataTable ReadExcelToTable(string path)//excel存放的路径
        {
            try
            {
                string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意
                using (OleDbConnection conn = new OleDbConnection(connstring))
                {
                    conn.Open();

                    System.Data.DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字

                    string firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字

                    string sql = string.Format("SELECT * FROM [{0}] WHERE [日期] is not null", firstSheetName); //查询字符串

                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);

                    DataSet set = new DataSet();

                    ada.Fill(set);

                    return set.Tables[0];
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();//首先根据打开文件对话框，选择excel表格
            fd.Filter = "表格|*.xls|所有文件(*.*)|*.*";//打开文件对话框筛选器
            string strPath;//文件完整的路径名
            if (fd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    strPath = fd.FileName;
                    string strCon = "provider=microsoft.ace.oledb.12.0;data source=" + strPath + ";extended properties='excel 12.0; HDR=NO;IMEX=1';";//关键是红色区域
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
        }//制表开始

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.Text == Convert.ToString(1))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }


                if (comboBox1.Text == Convert.ToString(2))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(3))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(4))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(5))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(6))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(7))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(8))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(9))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(10))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(11))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(12))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[13].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[13].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[13].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[13].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[13].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[13].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(13))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[14].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[14].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[14].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[14].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[14].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[14].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(14))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[15].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[15].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[15].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[15].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[15].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[15].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

                if (comboBox1.Text == Convert.ToString(15))
                {
                    double Jmax = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[16].Value) > Jmax)
                        {
                            Jmax = Convert.ToDouble(dataGridView1.Rows[i].Cells[16].Value);
                        }
                    }
                    textBox5.Text = Jmax.ToString("#0.00");//最大值

                    double Jmin = 10.0;
                    for (int i = 1; i <= dataGridView1.Rows.Count - 2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[16].Value) < Jmin)
                        {
                            Jmin = Convert.ToDouble(dataGridView1.Rows[i].Cells[16].Value);
                        }
                    }
                    textBox6.Text = Jmin.ToString("#0.00");//最小值

                    double Jsum = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum += Convert.ToDouble(dataGridView1.Rows[i].Cells[16].Value);
                    }
                    double Javg = Jsum / (dataGridView1.Rows.Count - 1);
                    textBox7.Text = Javg.ToString("#0.00");//平均值

                    double Jsum1 = 0;
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        Jsum1 += Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[16].Value) - Javg), 2);
                    }
                    double Jsqr = Math.Pow(Jsum1 / (dataGridView1.Rows.Count - 1), 0.5);
                    textBox8.Text = Jsqr.ToString("#0.00");//标准偏差

                    textBox9.Text = "±" + ((Jmax - Jmin) / 2);
                }

            }
            catch (Exception)
            {
                MessageBox.Show("请导入正确格式的数据");
            }

        }//选择单独显示的测点

        private void AnalysisOfData_Load(object sender, EventArgs e)//载入验证对象名
        {
            label1.Text = ("验证的对象是：" + FormBasicInformation.ObjName1);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open("E:\\4.PSV\\冷库验证项目模拟报告模板.doc");
            wordDoc.ActiveWindow.Visible = true;//打开word

            Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[6];//检索表格
            if (comboBox1.Text == Convert.ToString(1))
            {
                nowtable.Cell(2, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(2, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(2, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(2, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(2, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(2, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                ((Microsoft.Office.Interop.Word.Application)wordApp).Quit();
                wordApp = null;
            }


            if (comboBox1.Text == Convert.ToString(2))
            {

                nowtable.Cell(3, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(3, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(3, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(3, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(3, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(3, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(3))
            {

                nowtable.Cell(4, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(4, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(4, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(4, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(4, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(4, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(4))
            {

                nowtable.Cell(5, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(5, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(5, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(5, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(5, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(5, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(5))
            {

                nowtable.Cell(6, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(6, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(6, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(6, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(6, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(6, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(6))
            {

                nowtable.Cell(7, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(7, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(7, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(7, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(7, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(7, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(7))
            {

                nowtable.Cell(8, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(8, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(8, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(8, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(8, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(8, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(8))
            {

                nowtable.Cell(9, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(9, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(9, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(9, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(9, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(9, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(9))
            {

                nowtable.Cell(10, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(10, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(10, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(10, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(10, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(10, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(10))
            {

                nowtable.Cell(11, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(11, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(11, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(11, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(11, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(11, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(11))
            {

                nowtable.Cell(12, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(12, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(12, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(12, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(12, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(12, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }


            if (comboBox1.Text == Convert.ToString(12))
            {

                nowtable.Cell(13, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(13, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(13, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(13, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(13, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(13, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(13))
            {

                nowtable.Cell(14, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(14, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(14, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(14, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(14, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(14, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(14))
            {

                nowtable.Cell(15, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(15, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(15, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(15, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(15, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(15, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }

            if (comboBox1.Text == Convert.ToString(15))
            {

                nowtable.Cell(16, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                nowtable.Cell(16, 2).Range.InsertAfter(textBox2.Text);
                nowtable.Cell(16, 3).Range.InsertAfter(textBox3.Text);
                nowtable.Cell(16, 4).Range.InsertAfter(textBox4.Text);
                nowtable.Cell(16, 5).Range.InsertAfter(textBox5.Text);
                nowtable.Cell(16, 6).Range.InsertAfter(textBox6.Text);
                wordDoc.Save();
                wordApp.Quit();
                wordApp = null;
            }



        }//confirm
    }


}
