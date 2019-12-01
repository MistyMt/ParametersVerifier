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
            string filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName + ".doc";
            Spire.Doc.Document document = new Spire.Doc.Document(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox = document.TextBoxes[2];
            Spire.Doc.Documents.Paragraph paragraph = textBox.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange = paragraph.AppendText(textBox3.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox1 = document.TextBoxes[3];
            Spire.Doc.Documents.Paragraph paragraph1 = textBox1.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange1 = paragraph1.AppendText(textBox4.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox2 = document.TextBoxes[4];
            Spire.Doc.Documents.Paragraph paragraph2 = textBox2.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange2 = paragraph2.AppendText(textBox7.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox33 = document.TextBoxes[5];
            Spire.Doc.Documents.Paragraph paragraph33 = textBox33.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange33 = paragraph33.AppendText(textBox10.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox5 = document.TextBoxes[6];
            Spire.Doc.Documents.Paragraph paragraph5 = textBox5.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange5 = paragraph5.AppendText(textBox13.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox6 = document.TextBoxes[7];
            Spire.Doc.Documents.Paragraph paragraph6 = textBox6.Body.AddParagraph();
            Spire.Doc.Fields.TextRange textRange6 = paragraph6.AppendText(textBox16.Text);
            document.SaveToFile(filename, FileFormat.Docx);

            Spire.Doc.Fields.TextBox textBox8 = document.TextBoxes[8];
            Spire.Doc.Documents.Paragraph paragraph8 = textBox8.Body.AddParagraph();
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
    }
}
