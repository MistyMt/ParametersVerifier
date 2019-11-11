using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;


namespace Interface.Forms.DBForms
{
    public partial class selectedSensorData : Form
    {
        private string _senorName;
        public string senorName
        {
            set
            {
                _senorName = value;
            }
        }
        public selectedSensorData()
        {
            InitializeComponent();
        }

        private void selectedSensorData_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
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
                        comboBox1.DataSource = dt;
                        comboBox1.DisplayMember = "检测时间";

                        DataTable dts = sh.Select("SELECT " + " ID " + " FROM " + comboBox2.Text + ";");
                        comboBox3.DataSource = dts;
                        comboBox3.DisplayMember = "ID";
                    }
                    catch (Exception ex)
                    {
                        DataTable dt = new DataTable();
                        dt.Columns.Add("Error");
                        dt.Rows.Add(ex.ToString());
                        comboBox1.DataSource = dt;
                    }

                    conn.Close();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
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
                        DataTable dt = sh.Select("SELECT " + _senorName + " FROM " + comboBox2.Text + " WHERE ID=" + comboBox3.Text + ";");
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox3.SelectedIndex = comboBox1.SelectedIndex;
        }

        private void button3_Click(object sender, EventArgs e)
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
                        DataTable dt = sh.Select("SELECT " + _senorName + " FROM " + comboBox2.Text +";");
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

    }
}
