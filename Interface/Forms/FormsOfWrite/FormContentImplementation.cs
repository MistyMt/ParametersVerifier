using Spire.Doc;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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
            wordDoc.Close() ;
            wordApp.Quit();
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

            string ImagePath1 = pictureBox2.ImageLocation;
            string strKey1 = "6.1.2风机出风口布点";
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

            string ImagePath2 = pictureBox3.ImageLocation;
            string strKey2 = "6.1.3出入口布点";
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
                object objReplace2 = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
                if (wp2.Selection.Find.Execute(ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref objReplace2, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue))
                {
                    object Anchor2 = wp2.Selection.Range;
                    object LinkToFile2 = false;
                    object SaveWithDocument2 = true;
                    Microsoft.Office.Interop.Word.InlineShape Inlineshape2 = wp2.Selection.InlineShapes.AddPicture(
                        ImagePath2, ref LinkToFile2, ref SaveWithDocument2, ref Anchor2);
                    Inlineshape2.Select();
                    Microsoft.Office.Interop.Word.Shape shape2 = Inlineshape2.ConvertToShape();
                    shape2.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapBehind;

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

            string ImagePath3 = pictureBox4.ImageLocation;
            string strKey3 = "6.1.4死角布点";
            object MissingValue3 = Type.Missing;
            bool isFindSealLoc3 = false;
            Microsoft.Office.Interop.Word.Application wp3 = null;
            Microsoft.Office.Interop.Word.Document wd3 = null;
            try
            {
                wp3 = new Microsoft.Office.Interop.Word.Application();
                wd3 = wp3.Documents.Open(ref filename1, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue);
                wp3.Selection.Find.ClearFormatting();
                wp3.Selection.Find.Replacement.ClearFormatting();
                wp3.Selection.Find.Text = strKey3;
                object objReplace3 = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
                if (wp3.Selection.Find.Execute(ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref objReplace3, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue))
                {
                    object Anchor3 = wp3.Selection.Range;
                    object LinkToFile3 = false;
                    object SaveWithDocument3 = true;
                    Microsoft.Office.Interop.Word.InlineShape Inlineshape3 = wp3.Selection.InlineShapes.AddPicture(
                        ImagePath3, ref LinkToFile3, ref SaveWithDocument3, ref Anchor3);
                    Inlineshape3.Select();
                    Microsoft.Office.Interop.Word.Shape shape3 = Inlineshape3.ConvertToShape();
                    shape3.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapBehind;

                    isFindSealLoc3 = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (wd3 != null)
                {
                    wd3.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wd3);
                    wd3 = null;
                }
                if (wp3 != null)
                {
                    wp3.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wp3);
                    wp3 = null;
                }

            }

            string ImagePath4 = pictureBox5.ImageLocation;
            string strKey4 = "6.1.5死角布点";
            object MissingValue4 = Type.Missing;
            bool isFindSealLoc4 = false;
            Microsoft.Office.Interop.Word.Application wp4 = null;
            Microsoft.Office.Interop.Word.Document wd4 = null;
            try
            {
                wp4 = new Microsoft.Office.Interop.Word.Application();
                wd4 = wp4.Documents.Open(ref filename1, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue);
                wp4.Selection.Find.ClearFormatting();
                wp4.Selection.Find.Replacement.ClearFormatting();
                wp4.Selection.Find.Text = strKey4;
                object objReplace4 = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
                if (wp4.Selection.Find.Execute(ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref objReplace4, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue))
                {
                    object Anchor4 = wp4.Selection.Range;
                    object LinkToFile4 = false;
                    object SaveWithDocument4 = true;
                    Microsoft.Office.Interop.Word.InlineShape Inlineshape4 = wp4.Selection.InlineShapes.AddPicture(
                        ImagePath4, ref LinkToFile4, ref SaveWithDocument4, ref Anchor4);
                    Inlineshape4.Select();
                    Microsoft.Office.Interop.Word.Shape shape4 = Inlineshape4.ConvertToShape();
                    shape4.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapBehind;

                    isFindSealLoc4 = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (wd4 != null)
                {
                    wd4.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wd4);
                    wd4 = null;
                }
                if (wp4 != null)
                {
                    wp4.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wp4);
                    wp4 = null;
                }

            }

            string ImagePath5 = pictureBox6.ImageLocation;
            string strKey5 = "6.1.6货架布点";
            object MissingValue5 = Type.Missing;
            bool isFindSealLoc5 = false;
            Microsoft.Office.Interop.Word.Application wp5 = null;
            Microsoft.Office.Interop.Word.Document wd5 = null;
            try
            {
                wp5 = new Microsoft.Office.Interop.Word.Application();
                wd5 = wp5.Documents.Open(ref filename1, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue);
                wp5.Selection.Find.ClearFormatting();
                wp5.Selection.Find.Replacement.ClearFormatting();
                wp5.Selection.Find.Text = strKey5;
                object objReplace5 = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
                if (wp5.Selection.Find.Execute(ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref objReplace5, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue))
                {
                    object Anchor5 = wp5.Selection.Range;
                    object LinkToFile5 = false;
                    object SaveWithDocument5 = true;
                    Microsoft.Office.Interop.Word.InlineShape Inlineshape5 = wp5.Selection.InlineShapes.AddPicture(
                        ImagePath5, ref LinkToFile5, ref SaveWithDocument5, ref Anchor5);
                    Inlineshape5.Select();
                    Microsoft.Office.Interop.Word.Shape shape5 = Inlineshape5.ConvertToShape();
                    shape5.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapBehind;

                    isFindSealLoc5 = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (wd5 != null)
                {
                    wd5.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wd5);
                    wd5 = null;
                }
                if (wp5 != null)
                {
                    wp5.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wp5);
                    wp5 = null;
                }

            }

            string ImagePath6 = pictureBox7.ImageLocation;
            string strKey6 = "6.1.7监控系统探头及验证环境布点";
            object MissingValue6 = Type.Missing;
            bool isFindSealLoc6 = false;
            Microsoft.Office.Interop.Word.Application wp6 = null;
            Microsoft.Office.Interop.Word.Document wd6 = null;
            try
            {
                wp6 = new Microsoft.Office.Interop.Word.Application();
                wd6 = wp6.Documents.Open(ref filename1, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue,
                ref MissingValue, ref MissingValue);
                wp6.Selection.Find.ClearFormatting();
                wp6.Selection.Find.Replacement.ClearFormatting();
                wp6.Selection.Find.Text = strKey6;
                object objReplace6 = Microsoft.Office.Interop.Word.WdReplace.wdReplaceNone;
                if (wp6.Selection.Find.Execute(ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue,
                                           ref MissingValue, ref objReplace6, ref MissingValue,
                                           ref MissingValue, ref MissingValue, ref MissingValue))
                {
                    object Anchor6 = wp6.Selection.Range;
                    object LinkToFile6 = false;
                    object SaveWithDocument6 = true;
                    Microsoft.Office.Interop.Word.InlineShape Inlineshape6 = wp6.Selection.InlineShapes.AddPicture(
                        ImagePath6, ref LinkToFile6, ref SaveWithDocument6, ref Anchor6);
                    Inlineshape6.Select();
                    Microsoft.Office.Interop.Word.Shape shape6 = Inlineshape6.ConvertToShape();
                    shape6.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapBehind;

                    isFindSealLoc6 = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (wd6 != null)
                {
                    wd6.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wd6);
                    wd6 = null;
                }
                if (wp6 != null)
                {
                    wp6.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wp6);
                    wp6 = null;
                }
                MessageBox.Show("导入成功！");
            }


            #region 删除后台word占用。
            System.Diagnostics.Process myproc = new System.Diagnostics.Process();
            //得到所有打开的进程
            try
            {
                foreach (Process thisproc in Process.GetProcessesByName("WINWORD"))
                {
                    if (!thisproc.CloseMainWindow())
                    {
                        thisproc.Kill();
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("杀死" + "WINWORD" + "失败！");
            }
            #endregion


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

            MessageBox.Show("导入成功！");
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
                    this.pictureBox4.Load(pathname);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
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
                    this.pictureBox5.Load(pathname);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
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
                    this.pictureBox6.Load(pathname);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
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
                    this.pictureBox7.Load(pathname);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
