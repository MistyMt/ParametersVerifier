using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Interface
{
    public partial class FormVerificationTime : Form
    {
        public FormVerificationTime()
        {
            InitializeComponent();
            using (SQLiteConnection conn = new SQLiteConnection(config.DataSource))
            {
                using (SQLiteCommand cmd = new SQLiteCommand())
                {
                    conn.Open();
                    cmd.Connection = conn;

                    SQLiteHelper sh = new SQLiteHelper(cmd);


                    try
                    {

                        //读取行数
                        int numOfAll = QueryCounts("满载", conn) + QueryCounts("空载", conn) + QueryCounts("开门", conn) + QueryCounts("断电", conn);

                        DataTable dt = sh.Select("SELECT  检测时间  FROM 空载 WHERE ID = 1");
                        string strTime = dt.Rows[0][0].ToString();

                        DataTable dt2 = sh.Select("SELECT  检测时间  FROM 满载 WHERE ID = " + QueryCounts("满载", conn).ToString() + ";");
                        string endTime = dt2.Rows[0][0].ToString();

                        int numOfAll2 = QueryCounts("满载", conn);
                        //DataTable dts = new DataTable();
                        //dts.Column.Add("C1");
                        //dts.Column.Add("C2");

                        //DataRow dr = dt.NewRow();

                        dataGridView1.Rows.Add();

                        dataGridView1.Rows[0].Cells[0].Value = numOfAll;
                        dataGridView1.Rows[0].Cells[1].Value = strTime;
                        dataGridView1.Rows[0].Cells[2].Value = endTime;





                    }
                    catch (Exception ex)
                    {
                    }

                    conn.Close();
                }
            }

        }
        public int QueryCounts(string table, SQLiteConnection conn)
        {//查询SQLite表的行数
            string sql = string.Format("select count(*) from {0}", table);
            SQLiteCommand cmd = new SQLiteCommand(sql, conn);
            int n = Convert.ToInt32(cmd.ExecuteScalar());

            return n;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
