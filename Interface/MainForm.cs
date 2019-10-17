using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections;
using System.Runtime.InteropServices;
using System.Windows.Forms.DataVisualization.Charting;
using Microsoft.Office.Interop.Excel;
using System.IO;
using AnyCAD.Platform;
using AnyCAD.Basic;
using Microsoft.Office.Interop.Word;
using System.Drawing;


namespace Interface
{
    public partial class MainForm : Form
    {
        private AnyCAD.Presentation.RenderWindow3d m_RenderView = null;

        public AnyCAD.Presentation.RenderWindow3d RenderView
        {
            get { return m_RenderView; }
            set { m_RenderView = value; }
        }
        public MainForm()
        {
            InitializeComponent();
            label1.Text = "当前验证对象为：" + Global.objectName;
            //Global.mainForms.Add(this);

            var container = this.panel3;
            m_RenderView = new AnyCAD.Presentation.RenderWindow3d();
            m_RenderView.Size = container.ClientSize;
            m_RenderView.Dock = System.Windows.Forms.DockStyle.Fill;
            container.Controls.Add(m_RenderView);

        }

        private void 选择对象ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Global.basicInformationForms.Count != 0)
            {
                Global.basicInformationForms[0].ShowDialog();
            }
            else
            {
                FormBasicInformation basicInformationForm = new FormBasicInformation();
                Global.basicInformationForms.Add(basicInformationForm);
                basicInformationForm.ShowDialog();
                if (Global.basicInformationForms.Count >= 1)
                {
                    buttonChooseObject.Enabled = false;
                }
            }
        }

        public void button2_Click(object sender, EventArgs e)
        {
            if (Global.basicInformationForms.Count != 0)
            {
                Global.basicInformationForms[0].ShowDialog();
            }
            else
            {
                MessageBox.Show("请先选择验证对象");
            }

        }

        private void 传感器信息查看ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormDataView form3 = new FormDataView();
            form3.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (Global.dataViewForms.Count != 0)
            {
                Global.dataViewForms[0].ShowDialog();
            }
            else
            {
                FormDataView dataViewForm = new FormDataView();
                Global.dataViewForms.Add(dataViewForm);
                dataViewForm.ShowDialog();
            }
        }

        private void 传感器信息导入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Global.dataViewForms.Count != 0)
            {
                Global.dataViewForms[0].ShowDialog();
            }
            else
            {
                FormDataView dataViewForm = new FormDataView();
                Global.dataViewForms.Add(dataViewForm);
                dataViewForm.ShowDialog();
            }
        }

        private void butClear_Click(object sender, EventArgs e)
        {
            RenderView.ClearScene();
        }




        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            var t = new Form();
            t.TopLevel = false;
            t.Parent = panel2;
            t.Show();
        }

        private void butMove_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("MoveNode");
        }

        private void 新建ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MainForm mainform = new MainForm();
            mainform.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panelPrimitive.BringToFront();
            panelPrimitive.Visible = true;
            panelPrimitive.AutoScroll = true;
            panelPrimitive.AutoScrollMinSize = new Size(221, 554);

        }

        private void button5_Click(object sender, EventArgs e)
        {
            panelFeature.BringToFront();
            panelFeature.Visible = true;
            panelFeature.AutoScroll = true;
            panelFeature.AutoScrollMinSize = new Size(221, 554);
        }

        private void buttonPrimitive_Click(object sender, EventArgs e)
        {
            panelPrimitive.BringToFront();
            panelPrimitive.HorizontalScroll.Enabled = false;
            panelPrimitive.Visible = true;
            panelPrimitive.AutoScroll = true;
            panelPrimitive.AutoScrollMinSize = new Size(225, 554);
        }

        private void buttonFeature_Click(object sender, EventArgs e)
        {
            panelFeature.BringToFront();
            panelFeature.HorizontalScroll.Enabled = false;
            panelFeature.Visible = true;
            panelFeature.AutoScroll = true;
            panelFeature.AutoScrollMinSize = new Size(221, 554);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            RenderView.ClearScene();

        }

        private void button16_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("MoveNode");
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Image Files (*.png;*.jpg;*.bmp)|*.png;*.jpg;*.bmp|All Files(*.*)|*.*";

            if (DialogResult.OK != dlg.ShowDialog())
                return;

            ImageNode node = new ImageNode();
            node.SetImage(new AnyCAD.Platform.Path(dlg.FileName));
            node.SetWidth(100);
            node.SetHeight(200);

            RenderView.ShowSceneNode(node);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            String fontName = "simhei.ttf";
            AnyCAD.Platform.Font font = GlobalInstance.FontManager.FindFont(fontName);
            if (font != null)
            {
                font.SetCharHeight(24);
                font.SetCharWdith(24);
                font.Update();
            }

            {
                Vector3 position = new Vector3(200, 200, 200);


                TextNode text = new TextNode();
                text.SetFontName(fontName);
                text.SetPosition(position);
                text.SetText("Text 定点文字功能示范");
                text.SetTextColor(new ColorValue(1, 0, 0, 1));
                text.SetOffset(new Vector2(-24 * 7, 0));
                RenderView.SceneManager.AddNode(text);

                PointNode pn = new PointNode();
                pn.SetPoint(position);
                RenderView.ShowSceneNode(pn);
            }
            {
                TextNode text = new TextNode();
                text.SetPosition(new Vector3(10, 50, 0));
                text.SetText("Text 固定文字功能示范");
                text.SetTextColor(new ColorValue(255, 255, 255, 3));
                RenderView.SceneManager.AddNode2d(text);
            }

            RenderView.RequestDraw();
        }

        private void button10_Click(object sender, EventArgs e)
        {

            {
                TopoShape box = GlobalInstance.BrepTools.MakeBox(new Vector3(40, -20, 0), new Vector3(0, 0, 1), new Vector3(30, 40, 60));

                SceneNode sceneNode = RenderView.ShowGeometry(box, 100);

                FaceStyle style = new FaceStyle();
                style.SetColor(new ColorValue(0.5f, 0.3f, 0, 0.5f));

                Texture texture = new Texture();
                texture.SetName("mytexture2");
                texture.SetFilePath(new AnyCAD.Platform.Path("E:\\198.png"));
                style.SetTexture(0, texture);

                style.SetTransparent(true);
                sceneNode.SetFaceStyle(style);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            TopoShape sphere = GlobalInstance.BrepTools.MakeSphere(new Vector3(0, 0, 40), 40);
            SceneNode node = RenderView.ShowGeometry(sphere, 100);


            Texture texture = new Texture();
            texture.SetName("mytexture3");
            texture.SetFilePath(new AnyCAD.Platform.Path("#test\\land_ocean_ice_2048.jpg"));

            FaceStyle style = new FaceStyle();
            style.SetTexture(0, texture);

            node.SetFaceStyle(style);
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            double len = 20 * Convert.ToDouble(10);
            double width = 20 * Convert.ToDouble(10);
            double height = 20 * Convert.ToDouble(5);

            TopoShape box = GlobalInstance.BrepTools.MakeBox(Vector3.ZERO, new Vector3(len, 0, 0), width, height);
            RenderableGeometry geom = new RenderableGeometry();
            geom.SetGeometry(box);
            geom.SetShapeFilter((int)EnumPickMode.RF_Edge);// only display face
            EntitySceneNode node = new EntitySceneNode();
            node.SetEntity(geom);
            RenderView.ShowSceneNode(node);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            TopoShape cylinder = GlobalInstance.BrepTools.MakeCylinder(new Vector3(180, 0, 0), new Vector3(0, 0, 1), 20, 100, 360);

            Matrix4 mat = GlobalInstance.MatrixBuilder.MakeRotation(90, Vector3.UNIT_X);
            cylinder = GlobalInstance.BrepTools.Transform(cylinder, mat);
            SceneNode sceneNode = RenderView.ShowGeometry(cylinder, 100);
            FaceStyle style = new FaceStyle();

            Texture texture = new Texture();
            texture.SetName("mytexture2");
            texture.SetFilePath(new AnyCAD.Platform.Path(@"E:\Depot\sdk.net.pro.2\weixin.jpg"));
            style.SetTexture(0, texture);

            sceneNode.SetFaceStyle(style);

            sceneNode.SetPickable(false);
        }
        DrawLineEditor drawLine;
        private void button6_Click(object sender, EventArgs e)
        {

            if (drawLine == null)
                drawLine = new DrawLineEditor();
            RenderView.ActiveEditor(drawLine);
        }

        private void button11_Click(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            TopoShape spiralCurve = GlobalInstance.BrepTools.MakeSpiralCurve(100, 10, 10, Coordinate3.UNIT_XYZ);
            RenderView.ShowGeometry(spiralCurve, 100);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            AxesWidget axes = new AxesWidget();
            Matrix4 trf = GlobalInstance.MatrixBuilder.MakeTranslate(100, 100, 100);
            axes.SetTransform(trf);

            RenderView.ShowSceneNode(axes);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ArrowWidget arrow = new ArrowWidget();
            Matrix4 trf = GlobalInstance.MatrixBuilder.MakeTranslate(100, 100, 100);
            arrow.SetTransform(trf);

            RenderView.ShowSceneNode(arrow);
        }

        private void shadeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.SetDisplayMode((int)(EnumDisplayStyle.DS_Face | EnumDisplayStyle.DS_Realistic));
        }

        private void shadeWithEdgeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.SetDisplayMode((int)(EnumDisplayStyle.DS_Face | EnumDisplayStyle.DS_Edge | EnumDisplayStyle.DS_Realistic));

        }

        private void edgeWithPointsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.SetDisplayMode((int)(EnumDisplayStyle.DS_Edge | EnumDisplayStyle.DS_Vertex));

        }

        private void edgeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.Renderer.SetDisplayMode((int)EnumDisplayStyle.DS_Edge);
            RenderView.RequestDraw();
        }

        private void orbitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("Orbit");

        }

        private void zoomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("Zoom");

        }

        private void panToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("Pan");
        }

        private void zoomAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.FitAll();
        }

        private void zoomByRectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("ZoomByRect");

        }
        private bool mShowGrid = true;
        private void showGridToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mShowGrid = !mShowGrid;
            RenderView.ShowWorkingGrid(mShowGrid);
        }

        private void grideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Vector3 center = new Vector3(0, 0, 20);
            WorkingPlane wp1 = RenderView.Renderer.GetWorkingPlane();
            wp1.SetPosition(center);
            RenderView.RequestDraw();
        }

        private void prespectiveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("ProjectionMode");
            RenderView.Renderer.SetSkyBox("Early Morning");

            RenderView.RequestDraw();
        }

        private void topToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.Renderer.SetStandardView(EnumStandardView.SV_Top);
            RenderView.RequestDraw();
        }

        private void frontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.Renderer.SetStandardView(EnumStandardView.SV_Front);
            RenderView.RequestDraw();
        }

        private void pickByClickToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("Pick");
        }

        private void pickByRectangleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("RectPick");
        }

        private void singlePickToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("PickClearMode", "SinglePick");
        }

        private void multiPickToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.ExecuteCommand("PickClearMode", "MultiPick");
        }

        private void pickNodeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.SetPickMode((int)(EnumPickMode.RF_SceneNode | EnumPickMode.RF_Face));

        }

        private void pickGroupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.SetPickMode((int)(EnumPickMode.RF_GroupSceneNode | EnumPickMode.RF_Face));

        }

        private void pickFaceEdgePointToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenderView.SetPickMode((int)(EnumPickMode.RF_Default));

        }

        private void button14_Click(object sender, EventArgs e)
        {
            
            //创建"温度云图数据.txt"文件
            string filedir = Environment.CurrentDirectory.ToString() + "\\bin";
            string fullFilename = filedir + "\\温度云图数据.txt";
            System.IO.File.Delete(fullFilename);


            string fname = "";
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Excel File Dialog";
            fdlg.InitialDirectory = @filedir + "\\bin";
            fdlg.Filter = "All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                fname = fdlg.FileName;
            }

            if (fname != "")
            {
            string newTxtPath2 = filedir + "\\温度云图数据.txt";
            StreamWriter sw2 = new StreamWriter(newTxtPath2, true, Encoding.Default);
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet2 = xlWorkbook.Sheets[2];
            Microsoft.Office.Interop.Excel.Range xRange2 = xlWorksheet2.get_Range("C2:Q17");
            int rowCount2 = xRange2.Rows.Count;
            for (int k = 2; k <= rowCount2; k++)
            {
                string lineNum = k.ToString();
                double LineC2 = xlWorksheet2.Range["C" + lineNum].Value;
                double LineD2 = xlWorksheet2.Range["D" + lineNum].Value;
                double LineE2 = xlWorksheet2.Range["E" + lineNum].Value;
                double LineF2 = xlWorksheet2.Range["F" + lineNum].Value;
                double LineG2 = xlWorksheet2.Range["G" + lineNum].Value;
                double LineH2 = xlWorksheet2.Range["H" + lineNum].Value;
                double LineI2 = xlWorksheet2.Range["I" + lineNum].Value;
                double LineJ2 = xlWorksheet2.Range["J" + lineNum].Value;
                double LineK2 = xlWorksheet2.Range["K" + lineNum].Value;
                double LineL2 = xlWorksheet2.Range["L" + lineNum].Value;
                double LineM2 = xlWorksheet2.Range["M" + lineNum].Value;
                double LineN2 = xlWorksheet2.Range["N" + lineNum].Value;
                double LineO2 = xlWorksheet2.Range["O" + lineNum].Value;
                double LineP2 = xlWorksheet2.Range["P" + lineNum].Value;
                double LineQ2 = xlWorksheet2.Range["Q" + lineNum].Value;
                sw2.WriteLine(LineC2 + "\t" + LineD2 + "\t" + LineE2 + "\t" + LineF2 + "\t" + LineG2 + "\t" + LineH2 + "\t" + LineI2 + "\t" + LineJ2 + "\t" + LineK2 + "\t" + LineL2 + "\t" + LineM2 + "\t" + LineN2 + "\t" + LineO2 + "\t" + LineP2 + "\t" + LineQ2 + "\t");
            }
            sw2.Close();
            MessageBox.Show("ok");

            xlWorkbook.Close();
            
            //由"温度云图数据.txt"显示图像
            double[] xdata = { -8.37, -7.87, -7.37, -6.87, -6.37, -5.87, -5.37, -4.87, -4.37, -3.87, -3.37, -2.87, -2.37, -1.87, -1.37 };
            double[] ydata = { 0, 0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0 ,7.5};
            double[,] results = new double[xdata.Length, ydata.Length];
            double minValue = 10000000;
            double maxValue = -10000000;
            StreamReader sr = new StreamReader(fullFilename , Encoding.Default);
            String line;
            int numLine = 0;
            while ((line = sr.ReadLine()) != null)
            {
                String[] items = line.Split('\t');
                if (items.Length < 1)
                    continue;

                for (int ii = 0; ii < items.Length; ++ii)
                {
                    try
                    {
                        double val = double.Parse(items[ii]);
                        results[ii, numLine] = val;

                        minValue = Math.Min(minValue, val);
                        maxValue = Math.Max(maxValue, val);
                    }
                    catch (Exception)
                    {
                    }

                }

                ++numLine;
            }
            double range = maxValue - minValue;
            const long MaxValue = 0xff0000;
            const long MinValue = 0xffff00;
            const double Range = MaxValue - MinValue;
            float[] positionBuffer = new float[xdata.Length * ydata.Length * 3];
            float[] normalBuffer = new float[positionBuffer.Length];
            float[] colorBuffer = new float[positionBuffer.Length / 3 * 4];

            for (int jj = 0, lenjj = ydata.Length; jj < lenjj; ++jj)
                for (int ii = 0, lenii = xdata.Length; ii < lenii; ++ii)
                {
                    int idx = jj * lenii + ii;
                    positionBuffer[idx * 3] = (float)xdata[ii] * 10;
                    positionBuffer[idx * 3 + 1] = (float)ydata[jj] * 10;
                    positionBuffer[idx * 3 + 2] = 0;
                    normalBuffer[idx * 3] = 0;
                    normalBuffer[idx * 3 + 1] = 0;
                    normalBuffer[idx * 3 + 2] = 1;
                    double rst = results[ii, jj];
                    double ratio = (rst - minValue) / range;
                    long rgb = (long)(Range * ratio) + MinValue;
                    long red = rgb >> 16 & 0xFF;
                    long green = rgb >> 8 & 0xFF;
                    long blue = rgb & 0xFF;
                    colorBuffer[idx * 4] = red / 255.0f;
                    colorBuffer[idx * 4 + 1] = green / 255.0f;
                    colorBuffer[idx * 4 + 2] = blue / 255.0f;
                    colorBuffer[idx * 4 + 3] = 1.0f;
                }
            int faceCount = (xdata.Length - 1) * (ydata.Length - 1) * 2;
            uint[] facets = new uint[faceCount * 3];
            int faceId = 0;
            for (uint jj = 0, lenjj = (uint)ydata.Length; jj < lenjj - 1; ++jj)
                for (uint ii = 0, lenii = (uint)xdata.Length; ii < lenii - 1; ++ii)
                {
                    uint a = jj * lenii + ii;
                    uint b = a + 1;
                    uint c = (jj + 1) * lenii + ii;
                    uint d = c + 1;
                    /*
                    c----------d
                    |          |
                    |          |
                    |          |
                    a----------b
                    */
                    facets[faceId * 3] = a;
                    facets[faceId * 3 + 1] = d;
                    facets[faceId * 3 + 2] = c;

                    ++faceId;
                    facets[faceId * 3] = a;
                    facets[faceId * 3 + 1] = b;
                    facets[faceId * 3 + 2] = d;

                    ++faceId;
                }
            AABox bbox = new AABox();
            bbox.MinPt = new Vector3(positionBuffer[0], positionBuffer[1], positionBuffer[3]);
            bbox.MaxPt = new Vector3(positionBuffer[positionBuffer.Length - 3], positionBuffer[positionBuffer.Length - 2], positionBuffer[positionBuffer.Length - 1]);
            var entity = GlobalInstance.TopoShapeConvert.CreateColoredFaceEntity(positionBuffer, facets, normalBuffer, colorBuffer, bbox);
            var node = new EntitySceneNode();
            node.SetEntity(entity);
            RenderView.ShowSceneNode(node);
            }
        }

        public void buttonChooseObject_Click(object sender, EventArgs e)
        {
            FormBasicInformation basicInformationForm = new FormBasicInformation();
            Global.basicInformationForms.Add(basicInformationForm);
            basicInformationForm.ShowDialog();
            if (Global.basicInformationForms.Count >= 1)
            {
                buttonChooseObject.Enabled = false;
            }
            //if (basicInformationForm)
            //{

            //}
        }

        private void MainForm_Activated(object sender, EventArgs e)
        {
            label1.Text = "当前对象为：" + Global.objectName;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void 传感器信息重置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Global.dataViewForms.Clear();

        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (Global.analysisOfDataForms.Count != 0)
            {
                Global.analysisOfDataForms[0].ShowDialog();
            }
            else
            {
                FormAnalysisOfData analysisOfDataForm = new FormAnalysisOfData();
                Global.analysisOfDataForms.Add(analysisOfDataForm);
                analysisOfDataForm.ShowDialog();
            }
        }

        private void 重置对象ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Global.basicInformationForms.Clear();
            buttonChooseObject.Enabled = true;
        }

        private void 冷库ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Global.templateName = "冷库验证项目模拟报告模板.doc";
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void 导出图像_Click(object sender, EventArgs e)
        {
            Bitmap bit1 = new Bitmap(panel3.Width, panel3.Height);
            panel3.DrawToBitmap(bit1,new System.Drawing.Rectangle(0,0,panel3.Width,panel3.Height));
            int border = (panel3.Width - panel3.ClientSize.Width) / 2;//边框宽度
            int caption = (panel3.Height - panel3.ClientSize.Height) - border;//标题栏高度
            Bitmap bit2 = bit1.Clone(new System.Drawing.Rectangle(border, caption, panel3.ClientSize.Width, panel3.ClientSize.Height), System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            bit1.Save("C:\\Users\\13128\\Desktop\\AAA.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);//包括标题栏和边框
            bit2.Save("C:\\Users\\13128\\Desktop\\BBB.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);//不包括标题栏和边框
            bit1.Dispose();
            bit2.Dispose();
        }


        private void button18_Click(object sender, EventArgs e)
        {
            Bitmap screenShot = Global.GetScreenCapture();
            screenShot.Save("C:\\Users\\13128\\Desktop\\AAA.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

        }

        private void button20_Click(object sender, EventArgs e)
        {
            TopoShape rect = GlobalInstance.BrepTools.MakeRectangle(100, 50, 10, new Coordinate3(new Vector3(1, 1, 1), new Vector3(1, 1, 1), new Vector3(1, 1, 1), new Vector3(1, 1, 1)));
            rect = GlobalInstance.BrepTools.MakeFace(rect);
            RenderableGeometry geom = new RenderableGeometry();
            geom.SetGeometry(rect);
            EntitySceneNode node = new EntitySceneNode();
            node.SetEntity(geom);
            m_RenderView.ShowSceneNode(node);
        }




    }
}
