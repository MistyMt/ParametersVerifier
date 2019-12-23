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

namespace Interface.Forms.FormsOfWrite2
{
    public partial class AnalysisData : Form
    {
        public AnalysisData()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string pathname = string.Empty;
            OpenFileDialog file = new OpenFileDialog();
            file.InitialDirectory = System.Windows.Forms.Application.StartupPath.ToString() + "\\bin\\图片";
            file.Filter = "所有文件(*.*)|*.*";
            file.ShowDialog();
            if (file.FileName != string.Empty)
            {
                try
                {
                    pathname = file.FileName;   //获得文件的绝对路径
                    this.pictureBox1.Load(pathname);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string pathname = string.Empty;
            OpenFileDialog file = new OpenFileDialog();
            file.InitialDirectory = System.Windows.Forms.Application.StartupPath.ToString() + "\\bin\\图片";
            file.Filter = "所有文件(*.*)|*.*";
            file.ShowDialog();
            if (file.FileName != string.Empty)
            {
                try
                {
                    pathname = file.FileName;   //获得文件的绝对路径
                    this.pictureBox2.Load(pathname);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            object filename1 = Application.StartupPath + "\\bin\\" + Global.templateName;
            string ImagePath = pictureBox1.ImageLocation;
            string strKey = "8.4.1空载温度曲线图";
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

            }

            string ImagePath1 = pictureBox2.ImageLocation;
            string strKey1 = "8.4.3满载温度曲线图";
            object MissingValue1 = Type.Missing;
            bool isFindSealLoc1 = false;
            Microsoft.Office.Interop.Word.Application wp1 = null;
            Microsoft.Office.Interop.Word.Document wd1 = null;
            try
            {
                wp1 = new Microsoft.Office.Interop.Word.Application();
                wd1 = wp1.Documents.Open(ref filename1, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue);
                wp1.Selection.Find.ClearFormatting();
                wp1.Selection.Find.Replacement.ClearFormatting();
                wp1.Selection.Find.Text = strKey1;
                object objReplace1 = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
                if (wp1.Selection.Find.Execute(ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref objReplace1, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue))
                {
                    object Anchor1 = wp1.Selection.Range;
                    object LinkToFile1 = false;
                    object SaveWithDocument1 = true;
                    Microsoft.Office.Interop.Word.InlineShape Inlineshape1 = wp1.Selection.InlineShapes.AddPicture(
                        ImagePath1, ref LinkToFile1, ref SaveWithDocument1, ref Anchor1);
                    Inlineshape1.Select();
                    Microsoft.Office.Interop.Word.Shape shape1 = Inlineshape1.ConvertToShape();
                    shape1.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapBehind;

                    isFindSealLoc1 = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (wd1 != null)
                {
                    wd1.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wd1);
                    wd1 = null;
                }
                if (wp1 != null)
                {
                    wp1.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wp1);
                    wp1 = null;
                }

            }

            {
                string ImagePath2 = pictureBox3.ImageLocation;
                string strKey2 = "8.4.5热穿透测试温度曲线图";
                object MissingValue2 = Type.Missing;
                bool isFindSealLoc2 = false;
                Microsoft.Office.Interop.Word.Application wp2 = null;
                Microsoft.Office.Interop.Word.Document wd2 = null;
                try
                {
                    wp2 = new Microsoft.Office.Interop.Word.Application();
                    wd2 = wp2.Documents.Open(ref filename1, ref MissingValue,
                    ref MissingValue, ref MissingValue,
                    ref MissingValue, ref MissingValue,
                    ref MissingValue, ref MissingValue,
                    ref MissingValue, ref MissingValue,
                    ref MissingValue, ref MissingValue,
                    ref MissingValue, ref MissingValue,
                    ref MissingValue, ref MissingValue);
                    wp2.Selection.Find.ClearFormatting();
                    wp2.Selection.Find.Replacement.ClearFormatting();
                    wp2.Selection.Find.Text = strKey2;
                    object objReplace1 = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
                    if (wp2.Selection.Find.Execute(ref MissingValue, ref MissingValue, ref MissingValue,
                                               ref MissingValue, ref MissingValue, ref MissingValue,
                                               ref MissingValue, ref MissingValue, ref MissingValue,
                                               ref MissingValue, ref objReplace1, ref MissingValue,
                                               ref MissingValue, ref MissingValue, ref MissingValue))
                    {
                        object Anchor1 = wp2.Selection.Range;
                        object LinkToFile1 = false;
                        object SaveWithDocument1 = true;
                        Microsoft.Office.Interop.Word.InlineShape Inlineshape1 = wp2.Selection.InlineShapes.AddPicture(
                            ImagePath2, ref LinkToFile1, ref SaveWithDocument1, ref Anchor1);
                        Inlineshape1.Select();
                        Microsoft.Office.Interop.Word.Shape shape1 = Inlineshape1.ConvertToShape();
                        shape1.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapBehind;

                        isFindSealLoc2 = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    if (wd2 != null)
                    {
                        wd2.Close();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wd2);
                        wd2 = null;
                    }
                    if (wp2 != null)
                    {
                        wp2.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wp2);
                        wp2 = null;
                    }

                }
            }


            string filename = Application.StartupPath.ToString() + "\\bin\\" + Global.templateName + ".doc";
            Spire.Doc.Document document = new Spire.Doc.Document(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox = document.TextBoxes[1];
            Spire.Doc.Documents.Paragraph paragraph = textBox.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange = paragraph.AppendText(textBox1.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox111 = document.TextBoxes[2];
            Spire.Doc.Documents.Paragraph paragraph1 = textBox111.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange1 = paragraph1.AppendText(textBox2.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox1111 = document.TextBoxes[3];
            Spire.Doc.Documents.Paragraph paragraph11 = textBox1111.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange11 = paragraph11.AppendText(textBox3.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            MessageBox.Show("导入成功！");

            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string pathname = string.Empty;
            OpenFileDialog file = new OpenFileDialog();
            file.InitialDirectory = System.Windows.Forms.Application.StartupPath.ToString() + "\\bin\\图片";
            file.Filter = "所有文件(*.*)|*.*";
            file.ShowDialog();
            if (file.FileName != string.Empty)
            {
                try
                {
                    pathname = file.FileName;   //获得文件的绝对路径
                    this.pictureBox3.Load(pathname);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
