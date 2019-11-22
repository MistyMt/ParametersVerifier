using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Interface.Forms
{
    public partial class GeneralInfo : Form
    {
        public GeneralInfo()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region 从Resource文件夹提取模板
            FileInfo f111 = new FileInfo(System.Windows.Forms.Application.StartupPath + "\\bin\\" + Global.templateName + ".doc");
            if (f111.Exists)
            {
                f111.Delete();
                FileInfo f222 = new FileInfo(System.Windows.Forms.Application.StartupPath + "\\bin\\" + "\\Resource\\" + Global.templateName + ".doc");
                f222.CopyTo(System.Windows.Forms.Application.StartupPath + "\\bin\\" + Global.templateName + ".doc");
            }
            else 
            {
                FileInfo f222 = new FileInfo(System.Windows.Forms.Application.StartupPath + "\\bin\\" + "\\Resource\\" + Global.templateName + ".doc");
                f222.CopyTo(System.Windows.Forms.Application.StartupPath + "\\bin\\" + Global.templateName + ".doc");
            }
            #endregion

            object filename = Application.StartupPath + "\\bin\\" + Global.templateName;

            object G_Missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = false;//打开word

            Microsoft.Office.Interop.Word.Range myRange = wordDoc.Range();

            Microsoft.Office.Interop.Word.Find f = myRange.Find;
            f.Text = "委托单位：";
            f.ClearFormatting();

            bool finded = f.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange = wordDoc.Range(myRange.End, myRange.End + 28);
            myRange.Text = textBox1.Text;

            Microsoft.Office.Interop.Word.Range myRange1 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f1 = myRange1.Find;
            f1.Text = "验证对象：";
            f1.ClearFormatting();

            bool finded1 = f1.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange1 = wordDoc.Range(myRange1.End, myRange1.End + 28);
            myRange1.Text = textBox2.Text;

            Microsoft.Office.Interop.Word.Range myRange2 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f2 = myRange2.Find;
            f2.Text = "验证地点：";
            f2.ClearFormatting();

            bool finded2 = f2.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange2 = wordDoc.Range(myRange2.End, myRange2.End + 24);
            myRange2.Text = textBox3.Text;

            Microsoft.Office.Interop.Word.Range myRange3 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f3 = myRange3.Find;
            f3.Text = "验证时间：";
            f3.ClearFormatting();

            bool finded3 = f3.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange3 = wordDoc.Range(myRange3.End, myRange3.End + 4);
            myRange3.Text = textBox4.Text;

            Microsoft.Office.Interop.Word.Range myRange4 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f4 = myRange4.Find;
            f4.Text = "年";
            f4.ClearFormatting();

            bool finded4 = f4.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange4 = wordDoc.Range(myRange4.End, myRange4.End + 2);
            myRange4.Text = textBox5.Text;

            Microsoft.Office.Interop.Word.Range myRange5 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f5 = myRange5.Find;
            f5.Text = "月";
            f5.ClearFormatting();

            bool finded5 = f5.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange5 = wordDoc.Range(myRange5.End, myRange5.End + 2);
            myRange5.Text = textBox6.Text;

            Microsoft.Office.Interop.Word.Range myRange6 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f6 = myRange6.Find;
            f6.Text = "至 ";
            f6.ClearFormatting();

            bool finded6 = f6.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange6 = wordDoc.Range(myRange6.End, myRange6.End + 4);
            myRange6.Text = textBox7.Text;

            Microsoft.Office.Interop.Word.Range myRange7 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f7 = myRange7.Find;
            f7.Text = "年 ";
            f7.ClearFormatting();

            bool finded7 = f7.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange7 = wordDoc.Range(myRange7.End, myRange7.End + 1);
            myRange7.Text = textBox8.Text;

            Microsoft.Office.Interop.Word.Range myRange8 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f8 = myRange8.Find;
            f8.Text = "月 ";
            f8.ClearFormatting();

            bool finded8 = f8.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange8 = wordDoc.Range(myRange8.End, myRange8.End + 1);
            myRange8.Text = textBox9.Text;

            Microsoft.Office.Interop.Word.Range myRange9 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f9 = myRange9.Find;
            f9.Text = "编写人员：";
            f9.ClearFormatting();

            bool finded9 = f9.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange9 = wordDoc.Range(myRange9.End, myRange9.End + 16);
            myRange9.Text = textBox10.Text;

            Microsoft.Office.Interop.Word.Range myRange10 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f10 = myRange10.Find;
            f10.Text = "审核人员：";
            f10.ClearFormatting();

            bool finded10 = f10.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange10 = wordDoc.Range(myRange10.End, myRange10.End + 14);
            myRange10.Text = textBox11.Text;

            Microsoft.Office.Interop.Word.Range myRange11 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f11 = myRange11.Find;
            f11.Text = "批准人员：";
            f11.ClearFormatting();

            bool finded11 = f11.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange11 = wordDoc.Range(myRange11.End, myRange11.End + 14);
            myRange11.Text = textBox12.Text;

            wordDoc.Save();
            wordApp.Quit();
            wordApp = null;

            this.Close();
        }
        private void ClearAllText()
        {
            foreach (var ctr in Controls)
            {
                if (ctr is TextBox)
                {
                    ((TextBox)ctr).Text = "";
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ClearAllText();
        }

    }
}
