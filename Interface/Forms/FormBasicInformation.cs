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
    public partial class FormBasicInformation : Form
    {
        //private MainForm mainform;
        public FormBasicInformation()
        {
            InitializeComponent();


        }
        private static string ObjName;
        public static string ObjName1
        {
            get { return FormBasicInformation.ObjName; }
            set { FormBasicInformation.ObjName = value; }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            this.panel1.Visible = true;
            this.panel1.BringToFront();
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (comboBox1.Text == "冷库")
            {
                this.panel1.BringToFront();
                this.panel1.Visible = true;
                this.panel2.Visible = false;
                this.panel3.Visible = false;
            }
            else if (comboBox1.Text == "高温热处理炉")
            {
                this.panel2.BringToFront();
                this.panel2.Visible = true;
                this.panel1.Visible = false;
                this.panel3.Visible = false;
            }
            else if (comboBox1.Text == "灭菌器")
            {
                this.panel3.BringToFront();
                this.panel3.Visible = true;
                this.panel1.Visible = false;
                this.panel2.Visible = false;
            }
        }//切换panel
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

        private void button2_Click(object sender, EventArgs e)
        {
            ClearAllText();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Global.templateName = comboBox2.Text;
            object filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName;

            object G_Missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = true;//打开word

            Microsoft.Office.Interop.Word.Range myRange = wordDoc.Range();

            Microsoft.Office.Interop.Word.Find f = myRange.Find;
            f.Text = "冷库名称/编号：";
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
            f1.Text = "：长";
            f1.ClearFormatting();

            bool finded1 = f1.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange1 = wordDoc.Range(myRange1.End, myRange1.End + 4);
            myRange1.Text = textBox2.Text;

            Microsoft.Office.Interop.Word.Range myRange2 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f2 = myRange2.Find;
            f2.Text = "，宽";
            f2.ClearFormatting();

            bool finded2 = f2.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange2 = wordDoc.Range(myRange2.End, myRange2.End + 4);
            myRange2.Text = textBox4.Text;

            Microsoft.Office.Interop.Word.Range myRange3 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f3 = myRange3.Find;
            f3.Text = "，高";
            f3.ClearFormatting();

            bool finded3 = f3.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange3 = wordDoc.Range(myRange3.End, myRange3.End + 4);
            myRange3.Text = textBox5.Text;

            Microsoft.Office.Interop.Word.Range myRange4 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f4 = myRange4.Find;
            f4.Text = "体积约为";
            f4.ClearFormatting();

            bool finded4 = f4.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange4 = wordDoc.Range(myRange4.End, myRange4.End + 6);
            myRange4.Text = textBox6.Text;

            Microsoft.Office.Interop.Word.Range myRange5 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f5 = myRange5.Find;
            f5.Text = "出入口";
            f5.ClearFormatting();

            bool finded5 = f5.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange5 = wordDoc.Range(myRange5.End, myRange5.End + 4);
            myRange5.Text = textBox3.Text;

            Microsoft.Office.Interop.Word.Range myRange6 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f6 = myRange6.Find;
            f6.Text = "货架数量";
            f6.ClearFormatting();

            bool finded6 = f6.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange6 = wordDoc.Range(myRange6.End, myRange6.End + 5);
            myRange6.Text = textBox7.Text;

            Microsoft.Office.Interop.Word.Range myRange7 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f7 = myRange7.Find;
            f7.Text = "监控系统探头";
            f7.ClearFormatting();

            bool finded7 = f7.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange7 = wordDoc.Range(myRange7.End, myRange7.End + 6);
            myRange7.Text = textBox8.Text;

            Microsoft.Office.Interop.Word.Range myRange8 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f8 = myRange8.Find;
            f8.Text = "6风机";
            f8.ClearFormatting();

            bool finded8 = f8.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange8 = wordDoc.Range(myRange8.End, myRange8.End + 4);
            myRange8.Text = textBox9.Text;

            Microsoft.Office.Interop.Word.Range myRange9 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f9 = myRange9.Find;
            f9.Text = "（  ";
            f9.ClearFormatting();

            bool finded9 = f9.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange9 = wordDoc.Range(myRange9.End, myRange9.End + 2);
            myRange9.Text = textBox10.Text;

            Microsoft.Office.Interop.Word.Range myRange10 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f10 = myRange10.Find;
            f10.Text = "~";
            f10.ClearFormatting();

            bool finded10 = f10.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange10 = wordDoc.Range(myRange10.End, myRange10.End + 2);
            myRange10.Text = textBox12.Text;

            Microsoft.Office.Interop.Word.Range myRange11 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f11 = myRange11.Find;
            f11.Text = "（   ";
            f11.ClearFormatting();

            bool finded11 = f11.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange11 = wordDoc.Range(myRange11.End, myRange11.End + 2);
            myRange11.Text = textBox11.Text;

            Microsoft.Office.Interop.Word.Range myRange12 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f12 = myRange12.Find;
            f12.Text = "~ ";
            f12.ClearFormatting();

            bool finded12 = f12.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange12 = wordDoc.Range(myRange12.End, myRange12.End + 2);
            myRange12.Text = textBox13.Text;

            //wordDoc.Close();

            Global.objectName = comboBox1.Text;
            Global.objectLen = Convert.ToInt32( textBox2.Text);
            Global.objectWidth = Convert.ToInt32(textBox4.Text);
            Global.objectHeight = Convert.ToInt32(textBox5.Text);

            this.Close();
        }
    }
}
