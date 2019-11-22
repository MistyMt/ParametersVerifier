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
    public partial class FormOpenImpact : Form
    {
        public FormOpenImpact()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName + ".doc";
            Spire.Doc.Document document = new Spire.Doc.Document(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox = document.TextBoxes[12];
            Spire.Doc.Documents.Paragraph paragraph = textBox.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange = paragraph.AppendText(textBox1.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            object filename1 = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName;
            string ImagePath = Environment.CurrentDirectory.ToString() + "\\bin\\" + "ObjectPicturePreView.png";
            string strKey = "7.6.2 开门测试期间温度变化曲线图";
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
            this.Close();
        }
    }
}
