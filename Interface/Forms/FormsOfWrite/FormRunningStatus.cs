using System;
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
    public partial class FormRunningStatus : Form
    {
        public FormRunningStatus()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = true;
            Microsoft.Office.Interop.Word.Table nowtable2 = wordDoc.Tables[9];
            nowtable2.Cell(3, 3).Range.InsertAfter(textBox1.Text);
            nowtable2.Cell(4, 3).Range.InsertAfter(textBox2.Text);
            nowtable2.Cell(5, 3).Range.InsertAfter(textBox3.Text);
            nowtable2.Cell(7, 3).Range.InsertAfter(textBox4.Text);
            nowtable2.Cell(8, 3).Range.InsertAfter(textBox5.Text);
            nowtable2.Cell(9, 3).Range.InsertAfter(textBox6.Text);
            nowtable2.Cell(2, 4).Range.InsertAfter(textBox7.Text);
        }
    }
}
