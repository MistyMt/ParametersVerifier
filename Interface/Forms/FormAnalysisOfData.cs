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
using System.Data.SQLite;

namespace Interface
{
    public partial class FormAnalysisOfData : Form
    {
        public string SqlString;

        

        public FormAnalysisOfData()
        {
            InitializeComponent();
           
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
            //OpenFileDialog fd = new OpenFileDialog();//首先根据打开文件对话框，选择excel表格
            //fd.Filter = "表格|*.xls|所有文件(*.*)|*.*";//打开文件对话框筛选器
            //string strPath;//文件完整的路径名
            //if (fd.ShowDialog() == DialogResult.OK)
            //{
                try
                {
                    //strPath = fd.FileName;
                    string strPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\bin\\" + Global.dataSourceName;
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

            //}
        }//制表开始

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                double[] s1 = new double[15];
                double[] s2 = new double[15];
                double[] s3 = new double[15];
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
                    s1[0] = Javg;
                    s2[0] = Jsqr;
                    s3[0] = (Jmax - Jmin) / 2;
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
                    s1[1] = Javg;
                    s2[1] = Jsqr;
                    s3[1] = (Jmax - Jmin) / 2;
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
                    s1[2] = Javg;
                    s2[2] = Jsqr;
                    s3[2] = (Jmax - Jmin) / 2;
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
                    s1[3] = Javg;
                    s2[3] = Jsqr;
                    s3[3] = (Jmax - Jmin) / 2;
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
                    s1[4] = Javg;
                    s2[4] = Jsqr;
                    s3[4] = (Jmax - Jmin) / 2;
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
                    s1[5] = Javg;
                    s2[5] = Jsqr;
                    s3[5] = (Jmax - Jmin) / 2;
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
                    s1[6] = Javg;
                    s2[6] = Jsqr;
                    s3[6] = (Jmax - Jmin) / 2;
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
                    s1[7] = Javg;
                    s2[7] = Jsqr;
                    s3[7] = (Jmax - Jmin) / 2;
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
                    s1[8] = Javg;
                    s2[8] = Jsqr;
                    s3[8] = (Jmax - Jmin) / 2;
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
                    s1[9] = Javg;
                    s2[9] = Jsqr;
                    s3[9] = (Jmax - Jmin) / 2;
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
                    s1[10] = Javg;
                    s2[10] = Jsqr;
                    s3[10] = (Jmax - Jmin) / 2;
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
                    s1[11] = Javg;
                    s2[11] = Jsqr;
                    s3[11] = (Jmax - Jmin) / 2;
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
                    s1[12] = Javg;
                    s2[12] = Jsqr;
                    s3[12] = (Jmax - Jmin) / 2;
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
                    s1[13] = Javg;
                    s2[13] = Jsqr;
                    s3[13] = (Jmax - Jmin) / 2;
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
                    s1[14] = Javg;
                    s2[14] = Jsqr;
                    s3[14] = (Jmax - Jmin) / 2;
                    for (int i = 0; i < s1.Length; i++)
                    {
                        if (s1[i] <= 3)
                        {
                            textBox1.Text = "合格";
                        }
                        else
                        {
                            textBox1.Text = "不合格";
                        }

                    }
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
            object filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = false;//打开word
            if (Convert.ToString(dataGridView1.Rows[1].Cells[0].Value) == "空载")
            {
                Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[6];//检索表格
                if (comboBox1.Text == Convert.ToString(1))
                {
                    nowtable.Cell(2, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(2, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(2, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(2, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(2, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(2, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    ((Microsoft.Office.Interop.Word.Application)wordApp).Quit();
                    wordApp = null;
                }


                if (comboBox1.Text == Convert.ToString(2))
                {

                    nowtable.Cell(3, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(3, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(3, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(3, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(3, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(3, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(3))
                {

                    nowtable.Cell(4, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(4, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(4, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(4, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(4, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(4, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(4))
                {

                    nowtable.Cell(5, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(5, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(5, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(5, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(5, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(5, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(5))
                {

                    nowtable.Cell(6, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(6, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(6, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(6, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(6, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(6, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(6))
                {

                    nowtable.Cell(7, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(7, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(7, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(7, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(7, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(7, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(7))
                {
                    Microsoft.Office.Interop.Word.Table nowtable4 = wordDoc.Tables[10];
                    nowtable4.Cell(2, 2).Range.InsertAfter(Convert.ToString(4.5));
                    nowtable4.Cell(2, 3).Range.InsertAfter(textBox7.Text);
                    nowtable4.Cell(2, 4).Range.InsertAfter(Convert.ToString(4.5 - Convert.ToDouble(textBox7.Text)));
                    nowtable.Cell(8, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(8, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(8, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(8, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(8, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(8, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(8))
                {

                    nowtable.Cell(9, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(9, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(9, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(9, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(9, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(9, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(9))
                {

                    nowtable.Cell(10, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(10, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(10, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(10, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(10, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(10, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(10))
                {

                    nowtable.Cell(11, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(11, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(11, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(11, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(11, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(11, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(11))
                {

                    nowtable.Cell(12, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(12, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(12, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(12, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(12, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(12, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }


                if (comboBox1.Text == Convert.ToString(12))
                {

                    nowtable.Cell(13, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(13, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(13, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(13, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(13, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(13, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(13))
                {

                    nowtable.Cell(14, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(14, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(14, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(14, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(14, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(14, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(14))
                {

                    nowtable.Cell(15, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(15, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(15, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(15, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(15, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(15, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(15))
                {

                    nowtable.Cell(16, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(16, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(16, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(16, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(16, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(16, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }
            }

            if (Convert.ToString(dataGridView1.Rows[1].Cells[0].Value) == "满载")
            {
                textBox1.Text = "";
                Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[7];//检索表格
                if (comboBox1.Text == Convert.ToString(1))
                {
                    nowtable.Cell(2, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(2, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(2, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(2, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(2, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(2, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    ((Microsoft.Office.Interop.Word.Application)wordApp).Quit();
                    wordApp = null;
                }


                if (comboBox1.Text == Convert.ToString(2))
                {

                    nowtable.Cell(3, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(3, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(3, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(3, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(3, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(3, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(3))
                {

                    nowtable.Cell(4, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(4, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(4, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(4, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(4, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(4, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(4))
                {

                    nowtable.Cell(5, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(5, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(5, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(5, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(5, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(5, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(5))
                {

                    nowtable.Cell(6, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(6, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(6, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(6, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(6, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(6, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(6))
                {

                    nowtable.Cell(7, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(7, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(7, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(7, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(7, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(7, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(7))
                {
                    Microsoft.Office.Interop.Word.Table nowtable4 = wordDoc.Tables[10];
                    nowtable4.Cell(3, 2).Range.InsertAfter(Convert.ToString(4.5));
                    nowtable4.Cell(3, 3).Range.InsertAfter(textBox7.Text);
                    nowtable4.Cell(3, 4).Range.InsertAfter(Convert.ToString(4.5 - Convert.ToDouble(textBox7.Text)));
                    nowtable.Cell(8, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(8, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(8, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(8, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(8, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(8, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(8))
                {

                    nowtable.Cell(9, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(9, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(9, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(9, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(9, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(9, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(9))
                {

                    nowtable.Cell(10, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(10, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(10, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(10, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(10, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(10, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(10))
                {

                    nowtable.Cell(11, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(11, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(11, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(11, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(11, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(11, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(11))
                {

                    nowtable.Cell(12, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(12, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(12, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(12, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(12, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(12, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }


                if (comboBox1.Text == Convert.ToString(12))
                {

                    nowtable.Cell(13, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(13, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(13, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(13, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(13, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(13, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(13))
                {

                    nowtable.Cell(14, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(14, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(14, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(14, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(14, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(14, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(14))
                {

                    nowtable.Cell(15, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(15, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(15, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(15, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(15, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(15, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }

                if (comboBox1.Text == Convert.ToString(15))
                {

                    nowtable.Cell(16, 1).Range.InsertAfter(comboBox1.Text);//填充表格    
                    nowtable.Cell(16, 2).Range.InsertAfter(textBox5.Text);
                    nowtable.Cell(16, 3).Range.InsertAfter(textBox6.Text);
                    nowtable.Cell(16, 4).Range.InsertAfter(textBox7.Text);
                    nowtable.Cell(16, 5).Range.InsertAfter(textBox8.Text);
                    nowtable.Cell(16, 6).Range.InsertAfter(textBox9.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                }
            }



        }//confirm

        private void button1_Click(object sender, EventArgs e)
        {
            object filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = false;
            int j;
            int k;
            if (Convert.ToString(dataGridView1.Rows[1].Cells[0].Value) == "空载")
            {
                j = 0;
                k = 0;
                Microsoft.Office.Interop.Word.Table nowtable2 = wordDoc.Tables[4];
                nowtable2.Cell(j + 2, k + 5).Range.InsertAfter(dataGridView1[1, 1].Value.ToString());
                nowtable2.Cell(j + 2, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 2].Value.ToString());
                nowtable2.Cell(j + 3, k + 5).Range.InsertAfter(dataGridView1[1, 2].Value.ToString());
                nowtable2.Cell(j + 3, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 3].Value.ToString());
                nowtable2.Cell(j + 4, k + 5).Range.InsertAfter(dataGridView1[1, 3].Value.ToString());
                nowtable2.Cell(j + 4, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 4].Value.ToString());

                Microsoft.Office.Interop.Word.Table nowtable3 = wordDoc.Tables[5];
                nowtable3.Cell(1, 2).Range.InsertAfter(comboBox1.Text);
                nowtable3.Cell(2, 2).Range.InsertAfter(textBox5.Text);
                nowtable3.Cell(3, 2).Range.InsertAfter(textBox6.Text);
                nowtable3.Cell(4, 2).Range.InsertAfter(textBox7.Text);
            }
            if (Convert.ToString(dataGridView1.Rows[1].Cells[0].Value) == "满载")
            {
                j = 3;
                k = 0;
                Microsoft.Office.Interop.Word.Table nowtable2 = wordDoc.Tables[4];
                nowtable2.Cell(j + 2, k + 5).Range.InsertAfter(dataGridView1[1, 1].Value.ToString());
                nowtable2.Cell(j + 2, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 2].Value.ToString());
                nowtable2.Cell(j + 3, k + 5).Range.InsertAfter(dataGridView1[1, 2].Value.ToString());
                nowtable2.Cell(j + 3, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 3].Value.ToString());
                nowtable2.Cell(j + 4, k + 5).Range.InsertAfter(dataGridView1[1, 3].Value.ToString());
                nowtable2.Cell(j + 4, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 4].Value.ToString());
            }
            if (Convert.ToString(dataGridView1.Rows[1].Cells[0].Value) == "开门")
            {
                j = 6;
                k = -1;
                Microsoft.Office.Interop.Word.Table nowtable2 = wordDoc.Tables[4];
                nowtable2.Cell(j + 2, k + 5).Range.InsertAfter(dataGridView1[1, 1].Value.ToString());
                nowtable2.Cell(j + 2, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 2].Value.ToString());
                nowtable2.Cell(j + 3, k + 5).Range.InsertAfter(dataGridView1[1, 2].Value.ToString());
                nowtable2.Cell(j + 3, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 3].Value.ToString());
                nowtable2.Cell(j + 4, k + 5).Range.InsertAfter(dataGridView1[1, 3].Value.ToString());
                nowtable2.Cell(j + 4, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 4].Value.ToString());

                Microsoft.Office.Interop.Word.Table nowtable5 = wordDoc.Tables[11];
                nowtable5.Cell(3, 2).Range.InsertAfter(dataGridView1[2, 0].Value.ToString());
                nowtable5.Cell(4, 2).Range.InsertAfter(dataGridView1[3, 0].Value.ToString());
                nowtable5.Cell(5, 2).Range.InsertAfter(dataGridView1[4, 0].Value.ToString());
                nowtable5.Cell(6, 2).Range.InsertAfter(dataGridView1[5, 0].Value.ToString());
                nowtable5.Cell(7, 2).Range.InsertAfter(dataGridView1[6, 0].Value.ToString());
                nowtable5.Cell(8, 2).Range.InsertAfter(dataGridView1[7, 0].Value.ToString());
                nowtable5.Cell(9, 2).Range.InsertAfter(dataGridView1[8, 0].Value.ToString());
                nowtable5.Cell(10, 2).Range.InsertAfter(dataGridView1[9, 0].Value.ToString());
                nowtable5.Cell(11, 2).Range.InsertAfter(dataGridView1[10, 0].Value.ToString());
                nowtable5.Cell(12, 2).Range.InsertAfter(dataGridView1[11, 0].Value.ToString());
                nowtable5.Cell(13, 2).Range.InsertAfter(dataGridView1[12, 0].Value.ToString());
                nowtable5.Cell(3, 3).Range.InsertAfter(dataGridView1[2, 1].Value.ToString());
                nowtable5.Cell(4, 3).Range.InsertAfter(dataGridView1[3, 1].Value.ToString());
                nowtable5.Cell(5, 3).Range.InsertAfter(dataGridView1[4, 1].Value.ToString());
                nowtable5.Cell(6, 3).Range.InsertAfter(dataGridView1[5, 1].Value.ToString());
                nowtable5.Cell(7, 3).Range.InsertAfter(dataGridView1[6, 1].Value.ToString());
                nowtable5.Cell(8, 3).Range.InsertAfter(dataGridView1[7, 1].Value.ToString());
                nowtable5.Cell(9, 3).Range.InsertAfter(dataGridView1[8, 1].Value.ToString());
                nowtable5.Cell(10, 3).Range.InsertAfter(dataGridView1[9, 1].Value.ToString());
                nowtable5.Cell(11, 3).Range.InsertAfter(dataGridView1[10, 1].Value.ToString());
                nowtable5.Cell(12, 3).Range.InsertAfter(dataGridView1[11, 1].Value.ToString());
                nowtable5.Cell(13, 3).Range.InsertAfter(dataGridView1[12, 1].Value.ToString());
                nowtable5.Cell(3, 4).Range.InsertAfter(dataGridView1[2, 2].Value.ToString());
                nowtable5.Cell(4, 4).Range.InsertAfter(dataGridView1[3, 2].Value.ToString());
                nowtable5.Cell(5, 4).Range.InsertAfter(dataGridView1[4, 2].Value.ToString());
                nowtable5.Cell(6, 4).Range.InsertAfter(dataGridView1[5, 2].Value.ToString());
                nowtable5.Cell(7, 4).Range.InsertAfter(dataGridView1[6, 2].Value.ToString());
                nowtable5.Cell(8, 4).Range.InsertAfter(dataGridView1[7, 2].Value.ToString());
                nowtable5.Cell(9, 4).Range.InsertAfter(dataGridView1[8, 2].Value.ToString());
                nowtable5.Cell(10, 4).Range.InsertAfter(dataGridView1[9, 2].Value.ToString());
                nowtable5.Cell(11, 4).Range.InsertAfter(dataGridView1[10, 2].Value.ToString());
                nowtable5.Cell(12, 4).Range.InsertAfter(dataGridView1[11, 2].Value.ToString());
                nowtable5.Cell(13, 4).Range.InsertAfter(dataGridView1[12, 2].Value.ToString());
                nowtable5.Cell(3, 5).Range.InsertAfter(dataGridView1[2, 3].Value.ToString());
                nowtable5.Cell(4, 5).Range.InsertAfter(dataGridView1[3, 3].Value.ToString());
                nowtable5.Cell(5, 5).Range.InsertAfter(dataGridView1[4, 3].Value.ToString());
                nowtable5.Cell(6, 5).Range.InsertAfter(dataGridView1[5, 3].Value.ToString());
                nowtable5.Cell(7, 5).Range.InsertAfter(dataGridView1[6, 3].Value.ToString());
                nowtable5.Cell(8, 5).Range.InsertAfter(dataGridView1[7, 3].Value.ToString());
                nowtable5.Cell(9, 5).Range.InsertAfter(dataGridView1[8, 3].Value.ToString());
                nowtable5.Cell(10, 5).Range.InsertAfter(dataGridView1[9, 3].Value.ToString());
                nowtable5.Cell(11, 5).Range.InsertAfter(dataGridView1[10, 3].Value.ToString());
                nowtable5.Cell(12, 5).Range.InsertAfter(dataGridView1[11, 3].Value.ToString());
                nowtable5.Cell(13, 5).Range.InsertAfter(dataGridView1[12, 3].Value.ToString());

                Microsoft.Office.Interop.Word.Table nowtable6 = wordDoc.Tables[13];
                nowtable6.Cell(3, 2).Range.InsertAfter(dataGridView1[2, 0].Value.ToString());
                nowtable6.Cell(4, 2).Range.InsertAfter(dataGridView1[3, 0].Value.ToString());
                nowtable6.Cell(5, 2).Range.InsertAfter(dataGridView1[4, 0].Value.ToString());
                nowtable6.Cell(6, 2).Range.InsertAfter(dataGridView1[5, 0].Value.ToString());
                nowtable6.Cell(7, 2).Range.InsertAfter(dataGridView1[6, 0].Value.ToString());
                nowtable6.Cell(8, 2).Range.InsertAfter(dataGridView1[7, 0].Value.ToString());
                nowtable6.Cell(9, 2).Range.InsertAfter(dataGridView1[8, 0].Value.ToString());
                nowtable6.Cell(10, 2).Range.InsertAfter(dataGridView1[9, 0].Value.ToString());
                nowtable6.Cell(11, 2).Range.InsertAfter(dataGridView1[10, 0].Value.ToString());
                nowtable6.Cell(12, 2).Range.InsertAfter(dataGridView1[11, 0].Value.ToString());
                nowtable6.Cell(13, 2).Range.InsertAfter(dataGridView1[12, 0].Value.ToString());
                nowtable6.Cell(3, 3).Range.InsertAfter(dataGridView1[2, 1].Value.ToString());
                nowtable6.Cell(4, 3).Range.InsertAfter(dataGridView1[3, 1].Value.ToString());
                nowtable6.Cell(5, 3).Range.InsertAfter(dataGridView1[4, 1].Value.ToString());
                nowtable6.Cell(6, 3).Range.InsertAfter(dataGridView1[5, 1].Value.ToString());
                nowtable6.Cell(7, 3).Range.InsertAfter(dataGridView1[6, 1].Value.ToString());
                nowtable6.Cell(8, 3).Range.InsertAfter(dataGridView1[7, 1].Value.ToString());
                nowtable6.Cell(9, 3).Range.InsertAfter(dataGridView1[8, 1].Value.ToString());
                nowtable6.Cell(10, 3).Range.InsertAfter(dataGridView1[9, 1].Value.ToString());
                nowtable6.Cell(11, 3).Range.InsertAfter(dataGridView1[10, 1].Value.ToString());
                nowtable6.Cell(12, 3).Range.InsertAfter(dataGridView1[11, 1].Value.ToString());
                nowtable6.Cell(13, 3).Range.InsertAfter(dataGridView1[12, 1].Value.ToString());
                nowtable6.Cell(3, 4).Range.InsertAfter(dataGridView1[2, 2].Value.ToString());
                nowtable6.Cell(4, 4).Range.InsertAfter(dataGridView1[3, 2].Value.ToString());
                nowtable6.Cell(5, 4).Range.InsertAfter(dataGridView1[4, 2].Value.ToString());
                nowtable6.Cell(6, 4).Range.InsertAfter(dataGridView1[5, 2].Value.ToString());
                nowtable6.Cell(7, 4).Range.InsertAfter(dataGridView1[6, 2].Value.ToString());
                nowtable6.Cell(8, 4).Range.InsertAfter(dataGridView1[7, 2].Value.ToString());
                nowtable6.Cell(9, 4).Range.InsertAfter(dataGridView1[8, 2].Value.ToString());
                nowtable6.Cell(10, 4).Range.InsertAfter(dataGridView1[9, 2].Value.ToString());
                nowtable6.Cell(11, 4).Range.InsertAfter(dataGridView1[10, 2].Value.ToString());
                nowtable6.Cell(12, 4).Range.InsertAfter(dataGridView1[11, 2].Value.ToString());
                nowtable6.Cell(13, 4).Range.InsertAfter(dataGridView1[12, 2].Value.ToString());
                nowtable6.Cell(3, 5).Range.InsertAfter(dataGridView1[2, 3].Value.ToString());
                nowtable6.Cell(4, 5).Range.InsertAfter(dataGridView1[3, 3].Value.ToString());
                nowtable6.Cell(5, 5).Range.InsertAfter(dataGridView1[4, 3].Value.ToString());
                nowtable6.Cell(6, 5).Range.InsertAfter(dataGridView1[5, 3].Value.ToString());
                nowtable6.Cell(7, 5).Range.InsertAfter(dataGridView1[6, 3].Value.ToString());
                nowtable6.Cell(8, 5).Range.InsertAfter(dataGridView1[7, 3].Value.ToString());
                nowtable6.Cell(9, 5).Range.InsertAfter(dataGridView1[8, 3].Value.ToString());
                nowtable6.Cell(10, 5).Range.InsertAfter(dataGridView1[9, 3].Value.ToString());
                nowtable6.Cell(11, 5).Range.InsertAfter(dataGridView1[10, 3].Value.ToString());
                nowtable6.Cell(12, 5).Range.InsertAfter(dataGridView1[11, 3].Value.ToString());
                nowtable6.Cell(13, 5).Range.InsertAfter(dataGridView1[12, 3].Value.ToString());
            }
            if (Convert.ToString(dataGridView1.Rows[1].Cells[0].Value) == "断电")
            {
                j = 9;
                k = -1;
                Microsoft.Office.Interop.Word.Table nowtable2 = wordDoc.Tables[4];
                nowtable2.Cell(j + 2, k + 5).Range.InsertAfter(dataGridView1[1, 1].Value.ToString());
                nowtable2.Cell(j + 2, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 2].Value.ToString());
                nowtable2.Cell(j + 3, k + 5).Range.InsertAfter(dataGridView1[1, 2].Value.ToString());
                nowtable2.Cell(j + 3, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 3].Value.ToString());
                nowtable2.Cell(j + 4, k + 5).Range.InsertAfter(dataGridView1[1, 3].Value.ToString());
                nowtable2.Cell(j + 4, k + 6).Range.InsertAfter(dataGridView1[1, dataGridView1.RowCount - 4].Value.ToString());

                Microsoft.Office.Interop.Word.Table nowtable5 = wordDoc.Tables[12];
                nowtable5.Cell(4, 2).Range.InsertAfter(dataGridView1[2, 0].Value.ToString());
                nowtable5.Cell(5, 2).Range.InsertAfter(dataGridView1[3, 0].Value.ToString());
                nowtable5.Cell(6, 2).Range.InsertAfter(dataGridView1[4, 0].Value.ToString());
                nowtable5.Cell(7, 2).Range.InsertAfter(dataGridView1[5, 0].Value.ToString());
                nowtable5.Cell(8, 2).Range.InsertAfter(dataGridView1[6, 0].Value.ToString());
                nowtable5.Cell(9, 2).Range.InsertAfter(dataGridView1[7, 0].Value.ToString());
                nowtable5.Cell(10, 2).Range.InsertAfter(dataGridView1[8, 0].Value.ToString());
                nowtable5.Cell(11, 2).Range.InsertAfter(dataGridView1[9, 0].Value.ToString());
                nowtable5.Cell(12, 2).Range.InsertAfter(dataGridView1[10, 0].Value.ToString());
                nowtable5.Cell(13, 2).Range.InsertAfter(dataGridView1[11, 0].Value.ToString());
                nowtable5.Cell(4, 3).Range.InsertAfter(dataGridView1[1, 1].Value.ToString() + "~" + dataGridView1[1, dataGridView1.Rows.Count - 2].Value.ToString());
                nowtable5.Cell(5, 3).Range.InsertAfter(dataGridView1[1, 1].Value.ToString() + "~" + dataGridView1[1, dataGridView1.Rows.Count - 2].Value.ToString());
                nowtable5.Cell(6, 3).Range.InsertAfter(dataGridView1[1, 1].Value.ToString() + "~" + dataGridView1[1, dataGridView1.Rows.Count - 2].Value.ToString());
                nowtable5.Cell(7, 3).Range.InsertAfter(dataGridView1[1, 1].Value.ToString() + "~" + dataGridView1[1, dataGridView1.Rows.Count - 2].Value.ToString());
                nowtable5.Cell(8, 3).Range.InsertAfter(dataGridView1[1, 1].Value.ToString() + "~" + dataGridView1[1, dataGridView1.Rows.Count - 2].Value.ToString());
                nowtable5.Cell(9, 3).Range.InsertAfter(dataGridView1[1, 1].Value.ToString() + "~" + dataGridView1[1, dataGridView1.Rows.Count - 2].Value.ToString());
                nowtable5.Cell(10, 3).Range.InsertAfter(dataGridView1[1, 1].Value.ToString() + "~" + dataGridView1[1, dataGridView1.Rows.Count - 2].Value.ToString());
                nowtable5.Cell(11, 3).Range.InsertAfter(dataGridView1[1, 1].Value.ToString() + "~" + dataGridView1[1, dataGridView1.Rows.Count - 2].Value.ToString());
                nowtable5.Cell(12, 3).Range.InsertAfter(dataGridView1[1, 1].Value.ToString() + "~" + dataGridView1[1, dataGridView1.Rows.Count - 2].Value.ToString());
                nowtable5.Cell(13, 3).Range.InsertAfter(dataGridView1[1, 1].Value.ToString() + "~" + dataGridView1[1, dataGridView1.Rows.Count - 2].Value.ToString());
                nowtable5.Cell(4, 4).Range.InsertAfter(Convert.ToString(Convert.ToDouble(dataGridView1[2, dataGridView1.Rows.Count - 1].Value) - Convert.ToDouble(dataGridView1[2, 1].Value)));
                nowtable5.Cell(5, 4).Range.InsertAfter(Convert.ToString(Convert.ToDouble(dataGridView1[3, dataGridView1.Rows.Count - 1].Value) - Convert.ToDouble(dataGridView1[3, 1].Value)));
                nowtable5.Cell(6, 4).Range.InsertAfter(Convert.ToString(Convert.ToDouble(dataGridView1[4, dataGridView1.Rows.Count - 1].Value) - Convert.ToDouble(dataGridView1[4, 1].Value)));
                nowtable5.Cell(7, 4).Range.InsertAfter(Convert.ToString(Convert.ToDouble(dataGridView1[5, dataGridView1.Rows.Count - 1].Value) - Convert.ToDouble(dataGridView1[5, 1].Value)));
                nowtable5.Cell(8, 4).Range.InsertAfter(Convert.ToString(Convert.ToDouble(dataGridView1[6, dataGridView1.Rows.Count - 1].Value) - Convert.ToDouble(dataGridView1[6, 1].Value)));
                nowtable5.Cell(9, 4).Range.InsertAfter(Convert.ToString(Convert.ToDouble(dataGridView1[7, dataGridView1.Rows.Count - 1].Value) - Convert.ToDouble(dataGridView1[7, 1].Value)));
                nowtable5.Cell(10, 4).Range.InsertAfter(Convert.ToString(Convert.ToDouble(dataGridView1[8, dataGridView1.Rows.Count - 1].Value) - Convert.ToDouble(dataGridView1[8, 1].Value)));
                nowtable5.Cell(11, 4).Range.InsertAfter(Convert.ToString(Convert.ToDouble(dataGridView1[9, dataGridView1.Rows.Count - 1].Value) - Convert.ToDouble(dataGridView1[9, 1].Value)));
                nowtable5.Cell(12, 4).Range.InsertAfter(Convert.ToString(Convert.ToDouble(dataGridView1[10, dataGridView1.Rows.Count - 1].Value) - Convert.ToDouble(dataGridView1[10, 1].Value)));
                nowtable5.Cell(13, 4).Range.InsertAfter(Convert.ToString(Convert.ToDouble(dataGridView1[11, dataGridView1.Rows.Count - 1].Value) - Convert.ToDouble(dataGridView1[11, 1].Value)));
            }
            wordDoc.Save();
            wordApp.Quit();
            wordApp = null;
            MessageBox.Show("导入成功！");

        }

        private void button8_Click(object sender, EventArgs e)
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
                    string strCon = "provider=microsoft.ace.oledb.12.0;data source=" + strPath + ";extended properties='excel 12.0; HDR=NO;IMEX=1';";//关键是红色区域
                    OleDbConnection Con = new OleDbConnection(strCon);//建立连接
                    Con.Open();

                    System.Data.DataTable sheetsName = Con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字

                    string firstSheetName = sheetsName.Rows[1][2].ToString(); //得到第一个sheet的名字

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

            //}
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //OpenFileDialog fd = new OpenFileDialog();//首先根据打开文件对话框，选择excel表格
            //fd.Filter = "表格|*.xls|所有文件(*.*)|*.*";//打开文件对话框筛选器
            //string strPath;//文件完整的路径名
            //if (fd.ShowDialog() == DialogResult.OK)
            //{
                try
                {
                    string strPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\bin\\" + Global.dataSourceName;
                    //strPath = fd.FileName;
                    string strCon = "provider=microsoft.ace.oledb.12.0;data source=" + strPath + ";extended properties='excel 12.0; HDR=NO;IMEX=1';";//关键是红色区域
                    OleDbConnection Con = new OleDbConnection(strCon);//建立连接
                    Con.Open();

                    System.Data.DataTable sheetsName = Con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字

                    string firstSheetName = sheetsName.Rows[2][2].ToString(); //得到第一个sheet的名字

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

            //}
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //OpenFileDialog fd = new OpenFileDialog();//首先根据打开文件对话框，选择excel表格
            //fd.Filter = "表格|*.xls|所有文件(*.*)|*.*";//打开文件对话框筛选器
            //string strPath;//文件完整的路径名
            //if (fd.ShowDialog() == DialogResult.OK)
            //{
                try
                {
                    string strPath = System.Windows.Forms.Application.StartupPath.ToString() + "\\bin\\" + Global.dataSourceName;
                    //strPath = fd.FileName;
                    string strCon = "provider=microsoft.ace.oledb.12.0;data source=" + strPath + ";extended properties='excel 12.0; HDR=NO;IMEX=1';";//关键是红色区域
                    OleDbConnection Con = new OleDbConnection(strCon);//建立连接
                    Con.Open();

                    System.Data.DataTable sheetsName = Con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字

                    string firstSheetName = sheetsName.Rows[3][2].ToString(); //得到第一个sheet的名字

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

            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = false;
            Microsoft.Office.Interop.Word.Table nowtable1 = wordDoc.Tables[14];//检索表格
            nowtable1.Cell(7, 2).Range.InsertAfter(textBox1.Text);//结论导入
            nowtable1.Cell(2, 2).Range.InsertAfter("合格");
            nowtable1.Cell(3, 2).Range.InsertAfter("合格");
            nowtable1.Cell(4, 2).Range.InsertAfter("合格");
            nowtable1.Cell(5, 2).Range.InsertAfter("合格");
            nowtable1.Cell(6, 2).Range.InsertAfter("合格");
            nowtable1.Cell(8, 2).Range.InsertAfter("合格");
            wordDoc.Save();
            wordApp.Quit();
            wordApp = null;
            MessageBox.Show("导入成功！");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource != null)
            {

                System.Data.DataTable dt = (System.Data.DataTable)dataGridView1.DataSource;

                dt.Rows.Clear();

                dataGridView1.DataSource = dt;

            }

            else
            {

                dataGridView1.Rows.Clear();

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
        }
    }


}
