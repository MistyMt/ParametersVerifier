using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Interface
{
    public partial class ObjectSelect : Form
    {
        public ObjectSelect()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != string.Empty)
            {
                Global.objectName = comboBox1.Text;
                switch (Global.objectName)
                {
                    case "冷库":
                        Global.templateName = "冷库验证项目模拟报告模板";
                        break;
                    case "灭菌器":
                        Global.templateName = "灭菌器验证项目模拟报告模板";
                        break;
                    case "高温热处理炉":
                        Global.templateName = "高温热处理炉验证项目模拟报告模板";
                        break;
                }
                this.Close();
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {
            int num = 0;
            long length = 0;
            DirectoryInfo TheFolder = new DirectoryInfo(System.Windows.Forms.Application.StartupPath + "\\bin");
            //遍历文件
            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {
                if (NextFile.Name.Contains("模板") || NextFile.Name.Contains("~$"))
                {
                    length += NextFile.Length;
                    NextFile.Delete();
                    num += 1;
                }
            }
            var value = CountSize(length);
            MessageBox.Show("删除了" + num.ToString() + "个" + value + "文件");
        }
        /// <summary>
        /// 计算文件大小函数(保留两位小数),Size为字节大小
        /// </summary>
        /// <param name="Size">初始文件大小</param>
        /// <returns></returns>
        public static string CountSize(long Size)
        {
            string m_strSize = "";
            long FactSize = 0;
            FactSize = Size;
            if (FactSize < 1024.00)
                m_strSize = FactSize.ToString("F2") + " Byte";
            else if (FactSize >= 1024.00 && FactSize < 1048576)
                m_strSize = (FactSize / 1024.00).ToString("F2") + " K";
            else if (FactSize >= 1048576 && FactSize < 1073741824)
                m_strSize = (FactSize / 1024.00 / 1024.00).ToString("F2") + " M";
            else if (FactSize >= 1073741824)
                m_strSize = (FactSize / 1024.00 / 1024.00 / 1024.00).ToString("F2") + " G";
            return m_strSize;
        }

        private void label3_Click(object sender, EventArgs e)
        {
            Process.GetCurrentProcess().Kill();
        }

        private void label2_MouseEnter(object sender, EventArgs e)
        {
            this.label2.Font = new Font("宋书", 9, FontStyle.Underline);
        }

        private void label2_MouseLeave(object sender, EventArgs e)
        {
            this.label2.Font = new Font("宋书", 9, FontStyle.Regular);
        }

        private void label3_MouseEnter(object sender, EventArgs e)
        {
            this.label3.Font = new Font("宋书", 9, FontStyle.Underline);
        }

        private void label3_MouseLeave(object sender, EventArgs e)
        {
            this.label3.Font = new Font("宋书", 9, FontStyle.Regular);

        }

    }
}
