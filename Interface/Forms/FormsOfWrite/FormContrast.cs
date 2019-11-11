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
    public partial class FormContrast : Form
    {
        public FormContrast()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object filename = Environment.CurrentDirectory.ToString() + "\\bin\\" + Global.templateName;

            object G_Missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordDoc = wordApp.Documents.Open(filename);
            wordDoc.ActiveWindow.Visible = true;//打开word

            Microsoft.Office.Interop.Word.Range myRange = wordDoc.Range();

            Microsoft.Office.Interop.Word.Find f = myRange.Find;
            f.Text = "7.3.4分析说明  ";
            f.ClearFormatting();

            bool finded = f.Execute(ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing
                                    );

            myRange = wordDoc.Range(myRange.End, myRange.End + 28);
            myRange.Text = textBox1.Text;
        }
    }
}
