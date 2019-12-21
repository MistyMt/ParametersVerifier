﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Interface.Forms.DBForms
{
    public partial class Charts : Form
    {
        public Charts()
        {
            InitializeComponent();
        }

        private void Charts_Load(object sender, EventArgs e)
        {
            #region 读取测点列表
            using (SQLiteConnection conn = new SQLiteConnection(config.DataSource))
            {
                using (SQLiteCommand cmd = new SQLiteCommand())
                {
                    conn.Open();
                    cmd.Connection = conn;
                    SQLiteHelper sh = new SQLiteHelper(cmd);


                    try
                    {
                        List<string> columnName = new List<string>();
                        string sql = "PRAGMA table_info([满载]);";

                        SQLiteCommand cmd2 = new SQLiteCommand(sql, conn);
                        System.Data.SQLite.SQLiteDataReader dr = cmd2.ExecuteReader();

                        while (dr.Read())
                        {
                            columnName.Add(dr[1].ToString());

                        }
                        dr.Close();
                        conn.Close();
                        var colName = new List<string>();
                        for (int i = 3; i < columnName.Count; i++)
                        {
                            checkedListBox1.Items.Add(columnName[i].ToString());
                        }

                    }
                    catch (Exception ex)
                    {
                    }
                    conn.Close();
                }
            }
            #endregion
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count == 1)
            {

                using (SQLiteConnection conn = new SQLiteConnection(config.DataSource))
                {
                    using (SQLiteCommand cmd = new SQLiteCommand())
                    {
                        conn.Open();
                        cmd.Connection = conn;
                        SQLiteHelper sh = new SQLiteHelper(cmd);
                        try
                        {
                            DataTable dt = sh.Select("SELECT " + " 检测时间 " + " FROM " + comboBox2.Text + ";");
                            comboBox1.DisplayMember = "检测时间";
                            comboBox1.DataSource = dt;




                            DataTable dts = sh.Select("SELECT " + " ID " + " FROM " + comboBox2.Text + ";");
                            comboBox3.DisplayMember = "ID";
                            comboBox3.DataSource = dts;

                        }
                        catch (Exception ex)
                        {

                        }

                        conn.Close();
                    }

                }
            }
            else
            {
                MessageBox.Show("请单选测点。");
            }
        }



        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            // 刷新text
            StringBuilder show = new StringBuilder("目前已选择测点：");
            foreach (var item in checkedListBox1.CheckedItems)
            {
                show.Append(item.ToString() + ",");
            }
            string str = show.ToString();
            str = str.Substring(0, str.Length - 1);
            textBox1.Text = str;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }

            checkedListBox1.ClearSelected();
            // 刷新text
            StringBuilder show = new StringBuilder("目前已选择测点：");
            textBox1.Text = show.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //清空datagridview1
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

            using (SQLiteConnection conn = new SQLiteConnection(config.DataSource))
            {
                using (SQLiteCommand cmd = new SQLiteCommand())
                {
                    conn.Open();
                    cmd.Connection = conn;

                    SQLiteHelper sh = new SQLiteHelper(cmd);


                    try
                    {
                        DataTable dt = sh.Select("SELECT " + checkedListBox1.CheckedItems[0].ToString() + " FROM " + comboBox2.Text + " WHERE ID= " + comboBox3.Text + ";");
                        dataGridView1.DataSource = dt;
                    }
                    catch (Exception ex)
                    {
                        DataTable dt = new DataTable();
                        dt.Columns.Add("Error");
                        dt.Rows.Add(ex.ToString());
                        dataGridView1.DataSource = dt;
                    }

                    conn.Close();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //清空datagridview1
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


            using (SQLiteConnection conn = new SQLiteConnection(config.DataSource))
            {
                using (SQLiteCommand cmd = new SQLiteCommand())
                {
                    conn.Open();
                    cmd.Connection = conn;

                    SQLiteHelper sh = new SQLiteHelper(cmd);


                    try
                    {
                        DataTable dt = sh.Select("SELECT " + checkedListBox1.CheckedItems[0].ToString() + ",检测时间" + " FROM " + comboBox2.Text + ";");
                        dataGridView1.DataSource = dt;
                    }
                    catch (Exception ex)
                    {
                        DataTable dt = new DataTable();
                        dt.Columns.Add("Error");
                        dt.Rows.Add(ex.ToString());
                        dataGridView1.DataSource = dt;
                    }

                    conn.Close();
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {

                this.chart1.Series.Clear();
                //设置X/Y样式
                chart1.ChartAreas[0].AxisY.Title = "测点温度";
                chart1.ChartAreas[0].AxisX.Title = "时间";
                chart1.ChartAreas[0].AxisX.LabelStyle.Angle = 0;
                chart1.ChartAreas[0].AxisY.LabelStyle.Angle = 0;
                chart1.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
                chart1.ChartAreas[0].AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;
                chart1.ChartAreas[0].AxisY.Interval = 1;
                //chart1.ChartAreas[0].AxisY.IntervalType=double;

                //chart1.ChartAreas[0].AxisX.Enabled = AxisEnabled.False;
                //chart1.ChartAreas[0].AxisY.Enabled = AxisEnabled.False;
                //chart1.Titles[0].Text = "";

                //设置char样式

                this.chart1.Series.Add("测点温度");

                chart1.Series["测点温度"].MarkerColor = Color.Black;//设置标志
                chart1.Series["测点温度"].MarkerSize = 3;
                chart1.Series["测点温度"].MarkerStyle = MarkerStyle.Square;
                chart1.Series["测点温度"].IsValueShownAsLabel = true;//是否显示值
                chart1.Series["测点温度"].ChartType = SeriesChartType.Spline;//设置显示样式
                chart1.Series["测点温度"].BorderWidth = 1;
                chart1.Series["测点温度"].Color = Color.Blue;
                chart1.Series["测点温度"].ToolTip = "测点" + checkedListBox1.CheckedItems[0].ToString() + comboBox2.Text + "时数据统计";
                chart1.Series["测点温度"].YValueType = ChartValueType.Double;


                List<string> time = new List<string>();
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    time.Add(dataGridView1.Rows[i].Cells["检测时间"].Value.ToString());
                }

                List<double> data = new List<double>();
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    data.Add(Convert.ToDouble(dataGridView1.Rows[i].Cells[checkedListBox1.CheckedItems[0].ToString()].Value));
                }



                int nowRow = 0;
                if (dataGridView1.Rows.Count > 1)
                {
                    while (nowRow < dataGridView1.Rows.Count - 1)
                    {
                        chart1.Series["测点温度"].Points.AddXY(time[nowRow], data[nowRow]);
                        nowRow++;
                    }
                }

                chart1.ChartAreas[0].AxisY.Minimum = data.Min() * 0.90;
                chart1.ChartAreas[0].AxisY.Maximum = data.Max() * 1.10;

                chart1.ChartAreas[0].Axes[1].LabelStyle.Format = "N0";

            }
            catch (Exception)
            {
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif|PnG Image|*.png|Wmf  Image|*.wmf";
            saveFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath.ToString() + "\\bin" + "\\图片";
            saveFileDialog1.FilterIndex = 0;
            saveFileDialog1.FileName = "统计图";


            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                chart1.SaveImage(saveFileDialog1.FileName, System.Drawing.Imaging.ImageFormat.Png);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                comboBox3.SelectedIndex = comboBox1.SelectedIndex;

            }
            catch (Exception)
            {
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }


    }
}
