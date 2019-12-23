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

namespace Interface.Forms.FormsOfWrite3
{
    public partial class ImplementationOfTestingContent : Form
    {
        public ImplementationOfTestingContent()
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

            myRange = wordDoc.Range(myRange.End, myRange.End + 10);
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
            myRange1 = wordDoc.Range(myRange1.End, myRange1.End + 13);
            myRange1.Text = textBox1.Text;


            Microsoft.Office.Interop.Word.Range myRange2 = wordDoc.Range();
            Microsoft.Office.Interop.Word.Find f2 = myRange2.Find;
            f2.Text = "检测状态：空载（半载或满载）";
            f2.ClearFormatting();

            bool finded2 = f2.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange2 = wordDoc.Range(myRange2.End - 9, myRange2.End);
            myRange2.Text = comboBox1.Text;


            wordDoc.Save();
            wordDoc.Close(ref G_Missing, ref G_Missing, ref G_Missing);
            wordApp.Quit(ref G_Missing, ref G_Missing, ref G_Missing);
            wordApp = null;

            string ImagePath = pictureBox1.ImageLocation;
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

            }

            string filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName + ".doc";
            Spire.Doc.Document document = new Spire.Doc.Document(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox = document.TextBoxes[0];
            Spire.Doc.Documents.Paragraph paragraph = textBox.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange = paragraph.AppendText(textBox3.Text);
            document.SaveToFile(filename, FileFormat.Docx);
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
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
    }
}
