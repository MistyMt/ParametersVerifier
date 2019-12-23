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
    public partial class FormBeforeCalibrationValue : Form
    {
        public FormBeforeCalibrationValue()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();


            int index = dataGridView1.Rows.Add(row);

            dataGridView1.Rows[index].Cells[0].Value = textBox1.Text;

            dataGridView1.Rows[index].Cells[1].Value = textBox2.Text;//手动输入数据到datagridview
            dataGridView1.Rows[index].Cells[2].Value = textBox3.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            object filename =Application.StartupPath.ToString() + "\\bin\\" + Global.templateName;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = false;//打开word

            Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[15];//检索表格

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    nowtable.Cell(i + 2, j + 2).Range.InsertAfter(dataGridView1[j, i].Value.ToString());//填充表格
                }
            }
            wordDoc.Save();
            wordApp.Quit();
            wordApp = null;
        }
    }
}
