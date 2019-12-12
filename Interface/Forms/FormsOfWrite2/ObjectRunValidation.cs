using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Interface.Forms.FormsOfWrite2
{
    public partial class ObjectRunValidation : Form
    {
        public ObjectRunValidation()
        {
            InitializeComponent();
        }
        private void ClearAllText()
        {
            foreach (var ctr in Controls)
            {
                if (ctr is TextBox)//考虑是文本框的话
                {
                    ((TextBox)ctr).Text = "";
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            object filename = Application.StartupPath + "\\bin\\" + Global.templateName;

            object G_Missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = false;//打开word

            Microsoft.Office.Interop.Word.Range myRange = wordDoc.Range();

            Microsoft.Office.Interop.Word.Find f = myRange.Find;
            f.Text = "环境温度：";
            f.ClearFormatting();

            bool finded = f.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange = wordDoc.Range(myRange.End, myRange.End + 16);
            myRange.Text = textBox1.Text;

            Microsoft.Office.Interop.Word.Range myRange1 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f1 = myRange1.Find;
            f1.Text = "环境湿度：";
            f1.ClearFormatting();

            bool finded1 = f1.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange1 = wordDoc.Range(myRange1.End, myRange1.End + 16);
            myRange1.Text = textBox2.Text;

            Microsoft.Office.Interop.Word.Range myRange2 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f2 = myRange2.Find;
            f2.Text = "结果：";
            f2.ClearFormatting();

            bool finded2 = f2.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange2 = wordDoc.Range(myRange2.End, myRange2.End + 49);
            myRange2.Text = textBox3.Text;

            Microsoft.Office.Interop.Word.Range myRange3 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f3 = myRange3.Find;
            f3.Text = "检查人.：";
            f3.ClearFormatting();

            bool finded3 = f3.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange3 = wordDoc.Range(myRange3.End, myRange3.End + 18);
            myRange3.Text = textBox4.Text;

            Microsoft.Office.Interop.Word.Range myRange4 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f4 = myRange4.Find;
            f4.Text = "检查日期.： ";
            f4.ClearFormatting();

            bool finded4 = f4.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange4 = wordDoc.Range(myRange4.End, myRange4.End + 21);
            myRange4.Text = textBox5.Text;

            Microsoft.Office.Interop.Word.Range myRange5 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f5 = myRange5.Find;
            f5.Text = "复核人.：";
            f5.ClearFormatting();

            bool finded5 = f5.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange5 = wordDoc.Range(myRange5.End, myRange5.End + 18);
            myRange5.Text = textBox6.Text;

            Microsoft.Office.Interop.Word.Range myRange6 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f6 = myRange6.Find;
            f6.Text = "复核日期.：";
            f6.ClearFormatting();

            bool finded6 = f6.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange6 = wordDoc.Range(myRange6.End, myRange6.End + 10);
            myRange6.Text = textBox7.Text;

            Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[7];
            nowtable.Cell(1, 2).Range.InsertAfter(textBox10.Text);
            nowtable.Cell(2, 2).Range.InsertAfter(textBox8.Text);
            nowtable.Cell(3, 2).Range.InsertAfter(textBox9.Text);
            nowtable.Cell(4, 1).Range.InsertAfter(textBox12.Text);
            nowtable.Cell(4, 2).Range.InsertAfter(textBox11.Text);

            wordDoc.Save();
            wordApp.Quit(null, null, null);
            wordApp = null;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ClearAllText();
        }
    }
}
