﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
    }
}
