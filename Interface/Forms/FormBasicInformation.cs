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
            switch (Global.objectName)
            {
                case "冷库":
                    this.panel1.BringToFront();
                    this.panel1.Visible = true;
                    this.panel2.Visible = false;
                    this.panel3.Visible = false;
                    comboBox2.SelectedIndex = 0;
                    break;
                case "灭菌器":
                    this.panel3.BringToFront();
                    this.panel3.Visible = true;
                    this.panel1.Visible = false;
                    this.panel2.Visible = false;
                    comboBox2.SelectedIndex = 2;
                    break;
                case "高温热处理炉":
                    this.panel2.BringToFront();
                    this.panel2.Visible = true;
                    this.panel1.Visible = false;
                    this.panel3.Visible = false;
                    comboBox2.SelectedIndex = 1;
                    break;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "冷库验证项目模拟报告模板")
            {
                this.panel1.BringToFront();
                this.panel1.Visible = true;
                this.panel2.Visible = false;
                this.panel3.Visible = false;
            }
            else if (comboBox2.Text == "高温热处理炉验证项目模拟报告模板")
            {
                this.panel2.BringToFront();
                this.panel2.Visible = true;
                this.panel1.Visible = false;
                this.panel3.Visible = false;
            }
            else if (comboBox2.Text == "灭菌器验证项目模拟报告模板")
            {
                this.panel3.BringToFront();
                this.panel3.Visible = true;
                this.panel1.Visible = false;
                this.panel2.Visible = false;
            }
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

        private void button2_Click(object sender, EventArgs e)
        {
            ClearAllText();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "")
            {
                #region 冷库
                if (comboBox2.Text == "冷库验证项目模拟报告模板")
                {
                    object filename = Application.StartupPath + "\\bin\\" + Global.templateName;

                    object G_Missing = System.Reflection.Missing.Value;
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document wordDoc;
                    wordDoc = wordApp.Documents.Open(filename);
                    wordDoc.ActiveWindow.Visible = false;//打开word

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

                    wordDoc.Close();
                    wordApp.Quit();
                    wordApp = null;

                    try
                    {


                        Global.objectLen = Convert.ToDouble(textBox2.Text);
                        Global.objectWidth = Convert.ToDouble(textBox4.Text);
                        Global.objectHeight = Convert.ToDouble(textBox5.Text);


                        Global.tempLimitLow = Convert.ToDouble(textBox10.Text);
                        Global.tempLimitHigh = Convert.ToDouble(textBox12.Text);
                    }
                    catch (Exception)
                    {
                    }
                    this.Close();
                }
                #endregion

                #region 灭菌器
                if (comboBox2.Text == "灭菌器验证项目模拟报告模板")
                {
                    object filename = Application.StartupPath + "\\bin\\" + Global.templateName;
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document wordDoc;
                    wordDoc = wordApp.Documents.Open(filename);
                    wordDoc.ActiveWindow.Visible = false;
                    Microsoft.Office.Interop.Word.Table nowtable2 = wordDoc.Tables[4];
                    nowtable2.Cell(2, 3).Range.InsertAfter(textBox25.Text);
                    nowtable2.Cell(3, 3).Range.InsertAfter(textBox26.Text);
                    nowtable2.Cell(4, 3).Range.InsertAfter(textBox27.Text);
                    nowtable2.Cell(5, 3).Range.InsertAfter(textBox28.Text);
                    nowtable2.Cell(2, 4).Range.InsertAfter(textBox29.Text);
                    nowtable2.Cell(3, 4).Range.InsertAfter(textBox30.Text);
                    nowtable2.Cell(4, 4).Range.InsertAfter(textBox31.Text);
                    nowtable2.Cell(5, 4).Range.InsertAfter(textBox32.Text);
                    nowtable2.Cell(6, 1).Range.InsertAfter(textBox34.Text);
                    nowtable2.Cell(6, 2).Range.InsertAfter(textBox44.Text);
                    nowtable2.Cell(7, 1).Range.InsertAfter(textBox45.Text);
                    nowtable2.Cell(7, 2).Range.InsertAfter(textBox46.Text);

                    Microsoft.Office.Interop.Word.Table nowtable = wordDoc.Tables[5];
                    nowtable.Cell(2, 4).Range.InsertAfter(textBox33.Text);
                    nowtable.Cell(3, 4).Range.InsertAfter(textBox39.Text);
                    nowtable.Cell(4, 4).Range.InsertAfter(textBox41.Text);
                    nowtable.Cell(5, 4).Range.InsertAfter(textBox47.Text);
                    nowtable.Cell(6, 4).Range.InsertAfter(textBox48.Text);
                    nowtable.Cell(7, 1).Range.InsertAfter(textBox34.Text);
                    nowtable.Cell(7, 2).Range.InsertAfter(textBox44.Text);
                    nowtable.Cell(8, 1).Range.InsertAfter(textBox45.Text);
                    nowtable.Cell(8, 2).Range.InsertAfter(textBox46.Text);

                    Microsoft.Office.Interop.Word.Table nowtable1 = wordDoc.Tables[6];
                    nowtable1.Cell(2, 4).Range.InsertAfter(textBox35.Text);
                    nowtable1.Cell(3, 4).Range.InsertAfter(textBox36.Text);
                    nowtable1.Cell(4, 4).Range.InsertAfter(textBox37.Text);
                    nowtable1.Cell(5, 4).Range.InsertAfter(textBox38.Text);
                    nowtable1.Cell(6, 4).Range.InsertAfter(textBox40.Text);
                    nowtable1.Cell(7, 4).Range.InsertAfter(textBox42.Text);
                    nowtable1.Cell(8, 4).Range.InsertAfter(textBox43.Text);
                    nowtable1.Cell(9, 1).Range.InsertAfter(textBox34.Text);
                    nowtable1.Cell(9, 2).Range.InsertAfter(textBox44.Text);
                    nowtable1.Cell(10, 1).Range.InsertAfter(textBox45.Text);
                    nowtable1.Cell(10, 2).Range.InsertAfter(textBox46.Text);
                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                    this.Close();



                    try
                    {

                        Global.objectLen = Convert.ToDouble(textBox15.Text);
                        Global.objectWidth = Convert.ToDouble(textBox19.Text);
                        Global.objectHeight = Convert.ToDouble(textBox20.Text);



                        Global.tempLimitLow = Convert.ToDouble(textBox16.Text);

                        Global.tempLimitHigh = Convert.ToDouble(textBox18.Text);

                    }
                    catch (Exception)
                    {
                    }

                }

                #endregion

                #region 高温热处理炉
                if (comboBox2.Text == "高温热处理炉验证项目模拟报告模板")
                {
                    object filename = Application.StartupPath + "\\bin\\" + Global.templateName;

                    object G_Missing = System.Reflection.Missing.Value;
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document wordDoc;
                    wordDoc = wordApp.Documents.Open(filename);
                    wordDoc.ActiveWindow.Visible = false;//打开word

                    Microsoft.Office.Interop.Word.Range myRange = wordDoc.Range();

                    Microsoft.Office.Interop.Word.Find f = myRange.Find;
                    f.Text = "2.1.1热处理炉型号/编号：";
                    f.ClearFormatting();

                    bool finded = f.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange = wordDoc.Range(myRange.End, myRange.End + 28);
                    myRange.Text = textBox14.Text;

                    Microsoft.Office.Interop.Word.Range myRange1 = wordDoc.Range();
                    Microsoft.Office.Interop.Word.Find f1 = myRange1.Find;
                    f1.Text = "2.1.2热处理炉制造厂家：";
                    f1.ClearFormatting();

                    bool finded1 = f1.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange1 = wordDoc.Range(myRange1.End, myRange1.End + 28);
                    myRange1.Text = textBox23.Text;

                    Microsoft.Office.Interop.Word.Range myRange2 = wordDoc.Range();
                    Microsoft.Office.Interop.Word.Find f2 = myRange2.Find;
                    f2.Text = "2.1.3热处理炉规格（内径）：长";
                    f2.ClearFormatting();

                    bool finded2 = f2.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange2 = wordDoc.Range(myRange2.End, myRange2.End + 4);
                    myRange2.Text = textBox15.Text;

                    Microsoft.Office.Interop.Word.Range myRange3 = wordDoc.Range();
                    Microsoft.Office.Interop.Word.Find f3 = myRange3.Find;
                    f3.Text = "，宽";
                    f3.ClearFormatting();

                    bool finded3 = f3.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange3 = wordDoc.Range(myRange3.End, myRange3.End + 4);
                    myRange3.Text = textBox19.Text;

                    Microsoft.Office.Interop.Word.Range myRange4 = wordDoc.Range();
                    Microsoft.Office.Interop.Word.Find f4 = myRange4.Find;
                    f4.Text = "，高";
                    f4.ClearFormatting();

                    bool finded4 = f4.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange4 = wordDoc.Range(myRange4.End, myRange4.End + 4);
                    myRange4.Text = textBox20.Text;

                    Microsoft.Office.Interop.Word.Range myRange5 = wordDoc.Range();
                    Microsoft.Office.Interop.Word.Find f5 = myRange5.Find;
                    f5.Text = "体积约为";
                    f5.ClearFormatting();

                    bool finded5 = f5.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange5 = wordDoc.Range(myRange5.End, myRange5.End + 6);
                    myRange5.Text = textBox21.Text;

                    Microsoft.Office.Interop.Word.Range myRange6 = wordDoc.Range();
                    Microsoft.Office.Interop.Word.Find f6 = myRange6.Find;
                    f6.Text = "2.1.4监控系统探头";
                    f6.ClearFormatting();

                    bool finded6 = f6.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange6 = wordDoc.Range(myRange6.End, myRange6.End + 6);
                    myRange6.Text = textBox17.Text;

                    Microsoft.Office.Interop.Word.Range myRange7 = wordDoc.Range();
                    Microsoft.Office.Interop.Word.Find f7 = myRange7.Find;
                    f7.Text = "（  ";
                    f7.ClearFormatting();

                    bool finded7 = f7.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange7 = wordDoc.Range(myRange7.End, myRange7.End + 5);
                    myRange7.Text = textBox16.Text;

                    Microsoft.Office.Interop.Word.Range myRange8 = wordDoc.Range();
                    Microsoft.Office.Interop.Word.Find f8 = myRange8.Find;
                    f8.Text = "~";
                    f8.ClearFormatting();

                    bool finded8 = f8.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange8 = wordDoc.Range(myRange8.End, myRange8.End + 7);
                    myRange8.Text = textBox18.Text;

                    Microsoft.Office.Interop.Word.Range myRange9 = wordDoc.Range();
                    Microsoft.Office.Interop.Word.Find f9 = myRange9.Find;
                    f9.Text = "2.2.2保温精度范围：";
                    f9.ClearFormatting();

                    bool finded9 = f9.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange9 = wordDoc.Range(myRange9.End, myRange9.End + 7);
                    myRange9.Text = textBox22.Text;

                    Microsoft.Office.Interop.Word.Range myRange10 = wordDoc.Range();
                    Microsoft.Office.Interop.Word.Find f10 = myRange10.Find;
                    f10.Text = "2.2.3常用温度值：";
                    f10.ClearFormatting();

                    bool finded10 = f10.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                            ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                            );

                    myRange10 = wordDoc.Range(myRange10.End, myRange10.End + 12);
                    myRange10.Text = textBox24.Text;

                    wordDoc.Save();
                    wordApp.Quit();
                    wordApp = null;
                    this.Close();



                }
                #endregion
            }
            else
            {
                MessageBox.Show("请选择报表模板。");
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Global.objectHeight = Convert.ToDouble(textBox5.Text);
                textBox6.Text = ((Global.objectLen * Global.objectWidth * Global.objectHeight).ToString());
            }
            catch (Exception)
            {
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Global.objectWidth = Convert.ToDouble(textBox4.Text);
                textBox6.Text = ((Global.objectLen * Global.objectWidth * Global.objectHeight).ToString());
            }
            catch (Exception)
            {
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Global.objectLen = Convert.ToDouble(textBox2.Text);
                textBox6.Text = ((Global.objectLen * Global.objectWidth * Global.objectHeight).ToString());
            }
            catch (Exception)
            {
            }
        }

        private void label47_Click(object sender, EventArgs e)
        {

        }


    }
}
