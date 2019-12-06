using Spire.Doc;
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
    public partial class FormContentImplementation : Form
    {
        public FormContentImplementation()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object filename1 = Application.StartupPath + "\\bin\\" + Global.templateName;

            object G_Missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename1);
            wordDoc.ActiveWindow.Visible = false;//打开word

            Microsoft.Office.Interop.Word.Range myRange = wordDoc.Range();

            Microsoft.Office.Interop.Word.Find f = myRange.Find;
            f.Text = "布点数量：";
            f.ClearFormatting();

            bool finded = f.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange = wordDoc.Range(myRange.End, myRange.End + 6);
            myRange.Text = textBox1.Text;

            Microsoft.Office.Interop.Word.Range myRange1 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f1 = myRange1.Find;
            f1.Text = "仪表编号：";
            f1.ClearFormatting();

            bool finded1 = f1.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange1 = wordDoc.Range(myRange1.End, myRange1.End + 9);
            myRange1.Text = textBox2.Text;

            Microsoft.Office.Interop.Word.Range myRange2 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f2 = myRange2.Find;
            f2.Text = "测量点数量：";
            f2.ClearFormatting();

            bool finded2 = f2.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange2 = wordDoc.Range(myRange2.End, myRange2.End + 6);
            myRange2.Text = textBox6.Text;

            Microsoft.Office.Interop.Word.Range myRange3 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f3 = myRange3.Find;
            f3.Text = "仪表编号： ";
            f3.ClearFormatting();

            bool finded3 = f3.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange3 = wordDoc.Range(myRange3.End, myRange3.End + 8);
            myRange3.Text = textBox5.Text;

            Microsoft.Office.Interop.Word.Range myRange4 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f4 = myRange4.Find;
            f4.Text = "测量点数量： ";
            f4.ClearFormatting();

            bool finded4 = f4.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange4 = wordDoc.Range(myRange4.End, myRange4.End + 5);
            myRange4.Text = textBox9.Text;

            Microsoft.Office.Interop.Word.Range myRange5 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f5 = myRange5.Find;
            f5.Text = "仪表编号：  ";
            f5.ClearFormatting();

            bool finded5 = f5.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange5 = wordDoc.Range(myRange5.End, myRange5.End + 7);
            myRange5.Text = textBox8.Text;

            Microsoft.Office.Interop.Word.Range myRange6 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f6 = myRange6.Find;
            f6.Text = "测量点数量：  ";
            f6.ClearFormatting();

            bool finded6 = f6.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange6 = wordDoc.Range(myRange6.End, myRange6.End + 4);
            myRange6.Text = textBox12.Text;

            Microsoft.Office.Interop.Word.Range myRange7 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f7 = myRange7.Find;
            f7.Text = "仪表编号：   ";
            f7.ClearFormatting();

            bool finded7 = f7.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange7 = wordDoc.Range(myRange7.End, myRange7.End + 6);
            myRange7.Text = textBox11.Text;

            Microsoft.Office.Interop.Word.Range myRange8 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f8 = myRange8.Find;
            f8.Text = "测量点数量：   ";
            f8.ClearFormatting();

            bool finded8 = f8.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange8 = wordDoc.Range(myRange8.End, myRange8.End + 3);
            myRange8.Text = textBox15.Text;

            Microsoft.Office.Interop.Word.Range myRange9 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f9 = myRange9.Find;
            f9.Text = "仪表编号：    ";
            f9.ClearFormatting();

            bool finded9 = f9.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange9 = wordDoc.Range(myRange9.End, myRange9.End + 5);
            myRange9.Text = textBox14.Text;

            Microsoft.Office.Interop.Word.Range myRange10 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f10 = myRange10.Find;
            f10.Text = "测量点数量：    ";
            f10.ClearFormatting();

            bool finded10 = f10.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange10 = wordDoc.Range(myRange10.End, myRange10.End + 2);
            myRange10.Text = textBox18.Text;

            Microsoft.Office.Interop.Word.Range myRange11 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f11 = myRange11.Find;
            f11.Text = "仪表编号：     ";
            f11.ClearFormatting();

            bool finded11 = f11.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange11 = wordDoc.Range(myRange11.End, myRange11.End + 4);
            myRange11.Text = textBox17.Text;

            Microsoft.Office.Interop.Word.Range myRange12 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f12 = myRange12.Find;
            f12.Text = "测量点数量：     ";
            f12.ClearFormatting();

            bool finded12 = f12.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange12 = wordDoc.Range(myRange12.End, myRange12.End + 1);
            myRange12.Text = textBox21.Text;

            Microsoft.Office.Interop.Word.Range myRange13 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f13 = myRange13.Find;
            f13.Text = "仪表编号：      ";
            f13.ClearFormatting();

            bool finded13 = f13.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange13 = wordDoc.Range(myRange13.End, myRange13.End + 3);
            myRange13.Text = textBox20.Text;

            wordDoc.Save();
            wordApp.Quit();
            wordApp = null;

            string ImagePath = Environment.CurrentDirectory.ToString() + "\\bin\\" + "图片\\" + "ObjectPicturePreView.png";
            string strKey = "示意图：（上层，中层，下层）";
            object MissingValue = Type.Missing;
            bool isFindSealLoc = false;
            Microsoft.Office.Interop.Word.Application wp = null;
            Microsoft.Office.Interop.Word.Document wd = null;
            try
            {
                wp = new Microsoft.Office.Interop.Word.Application();
                wd = wp.Documents.Open(ref filename1, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue);
                wp.Selection.Find.ClearFormatting();
                wp.Selection.Find.Replacement.ClearFormatting();
                wp.Selection.Find.Text = strKey;
                object objReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
                if (wp.Selection.Find.Execute(ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref objReplace, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue))
                {
                    object Anchor = wp.Selection.Range;
                    object LinkToFile = false;
                    object SaveWithDocument = true;
                    Microsoft.Office.Interop.Word.InlineShape Inlineshape = wp.Selection.InlineShapes.AddPicture(
                        ImagePath, ref LinkToFile, ref SaveWithDocument, ref Anchor);
                    Inlineshape.Select();
                    Microsoft.Office.Interop.Word.Shape shape = Inlineshape.ConvertToShape();
                    shape.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapBehind;

                    isFindSealLoc = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (wd != null)
                {
                    wd.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wd);
                    wd = null;
                }
                if (wp != null)
                {
                    wp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wp);
                    wp = null;
                }
                MessageBox.Show("导入成功！");
            }

            string filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName + ".doc";
            Spire.Doc.Document document = new Spire.Doc.Document(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox = document.TextBoxes[2];
            Spire.Doc.Documents.Paragraph paragraph = textBox.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange = paragraph.AppendText(textBox3.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox111 = document.TextBoxes[3];
            Spire.Doc.Documents.Paragraph paragraph1 = textBox111.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange1 = paragraph1.AppendText(textBox4.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox22 = document.TextBoxes[4];
            Spire.Doc.Documents.Paragraph paragraph2 = textBox22.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange2 = paragraph2.AppendText(textBox7.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox33 = document.TextBoxes[5];
            Spire.Doc.Documents.Paragraph paragraph33 = textBox33.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange33 = paragraph33.AppendText(textBox10.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox55 = document.TextBoxes[6];
            Spire.Doc.Documents.Paragraph paragraph5 = textBox55.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange5 = paragraph5.AppendText(textBox13.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox66 = document.TextBoxes[7];
            Spire.Doc.Documents.Paragraph paragraph6 = textBox66.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange6 = paragraph6.AppendText(textBox16.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox88 = document.TextBoxes[8];
            Spire.Doc.Documents.Paragraph paragraph8 = textBox88.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange8 = paragraph8.AppendText(textBox19.Text);
            document.SaveToFile(filename, FileFormat.Docx);


            this.Close();

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
        private float X, Y;

        private void setTag(Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                con.Tag = con.Width + ":" + con.Height + ":" + con.Left + ":" + con.Top + ":" + con.Font.Size;
                if (con.Controls.Count > 0)
                    setTag(con);
            }
        }

        private void setControls(float newx, float newy, Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                string[] mytag = con.Tag.ToString().Split(new char[] { ':' });
                float a = Convert.ToSingle(mytag[0]) * newx;
                con.Width = (int)a;
                a = Convert.ToSingle(mytag[1]) * newy;
                con.Height = (int)(a);
                a = Convert.ToSingle(mytag[2]) * newx;
                con.Left = (int)(a);
                a = Convert.ToSingle(mytag[3]) * newy;
                con.Top = (int)(a);
                Single currentSize = Convert.ToSingle(mytag[4]) * newy;
                con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);
                if (con.Controls.Count > 0)
                {
                    setControls(newx, newy, con);
                }
            }
        }
        private void FormAfterCalibrationValue_Resize(object sender, EventArgs e)
        {
            float newx = (this.Width) / X;
            float newy = this.Height / Y;
            setControls(newx, newy, this);
            this.Text = this.Width.ToString() + " " + this.Height.ToString();
        }
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.Resize += new EventHandler(FormAfterCalibrationValue_Resize);
            X = this.Width;
            Y = this.Height;
            setTag(this);
        }

        private void FormContentImplementation_Load(object sender, EventArgs e)
        {
            var sList1 = new List<Sensor>();
            StringBuilder sb1 = new StringBuilder();
            var sList2 = new List<Sensor>();
            StringBuilder sb2 = new StringBuilder();
            var sList3 = new List<Sensor>();
            StringBuilder sb3 = new StringBuilder();
            var sList4 = new List<Sensor>();
            StringBuilder sb4 = new StringBuilder();
            var sList5 = new List<Sensor>();
            StringBuilder sb5 = new StringBuilder();
            var sList6 = new List<Sensor>();
            StringBuilder sb6 = new StringBuilder();
            foreach (Sensor vs in Global.sensors.Values)
            {
                switch (vs.positionName)
                {
                    case "均匀性布点":
                        sList1.Add(vs);
                        sb1.Append(vs.serialNumber + "\r\n");
                        break;
                    case "风机出风口布点":
                        sList2.Add(vs);
                        sb2.Append(vs.serialNumber + "\r\n");
                        break;
                    case "出入口布点":
                        sList3.Add(vs);
                        sb3.Append(vs.serialNumber + "\r\n");
                        break;
                    case "死角布点":
                        sList4.Add(vs);
                        sb4.Append(vs.serialNumber + "\r\n");
                        break;
                    case "货架布点":
                        sList5.Add(vs);
                        sb5.Append(vs.serialNumber + "\r\n");
                        break;
                    case "监控系统探头及验证环境布点":
                        sList6.Add(vs);
                        sb6.Append(vs.serialNumber + "\r\n");
                        break;
                }
            }
            textBox1.Text = sList1.Count.ToString();
            textBox2.Text = sb1.ToString();

            textBox6.Text = sList2.Count.ToString();
            textBox5.Text = sb2.ToString();

            textBox9.Text = sList3.Count.ToString();
            textBox8.Text = sb3.ToString();

            textBox12.Text = sList4.Count.ToString();
            textBox11.Text = sb4.ToString();

            textBox15.Text = sList4.Count.ToString();
            textBox14.Text = sb4.ToString();

            textBox18.Text = sList5.Count.ToString();
            textBox17.Text = sb5.ToString();

            textBox21.Text = sList6.Count.ToString();
            textBox20.Text = sb6.ToString();

        }
    }
}
