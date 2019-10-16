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
using AnyCAD.Platform;

namespace Interface
{
    public partial class FormDataView : Form
    {
        string[] x = new string[] { null, null, null, null, null, null, null, null, null, null };
        double[] y = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

        List<string> xList1 = new List<string>();
        List<double> yList1 = new List<double>();
        List<double> yList2 = new List<double>();
        List<double> yList3 = new List<double>();
        List<double> yList4 = new List<double>();
        List<double> yList5 = new List<double>();
        List<double> yList6 = new List<double>();
        List<double> yList7 = new List<double>();
        List<double> yList8 = new List<double>();
        List<double> yList9 = new List<double>();
        List<double> yList10 = new List<double>();


        public FormDataView()
        {
            InitializeComponent();
            var container = this.panel1;

            m_RenderView = new AnyCAD.Presentation.RenderWindow3d();

            m_RenderView.Size = container.ClientSize;
            m_RenderView.Dock = System.Windows.Forms.DockStyle.Fill;
            container.Controls.Add(m_RenderView);
        }//初始化

        private AnyCAD.Presentation.RenderWindow3d m_RenderView = null;//AnyCAD

        public AnyCAD.Presentation.RenderWindow3d RenderView
        {
            get { return m_RenderView; }
            set { m_RenderView = value; }
        }//AnyCAD

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                     chart1.Visible = true;
            string fname = "";
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Excel File Dialog";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";//文件名筛选器
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                fname = fdlg.FileName;
            }

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xRange = xlWorksheet.get_Range("A4:B784");
            Microsoft.Office.Interop.Excel.Range y1Range = xlApp.get_Range("B4:B784");
            Microsoft.Office.Interop.Excel.Range y2Range = xlApp.get_Range("C4:C784");
            Microsoft.Office.Interop.Excel.Range y3Range = xlApp.get_Range("D4:D784");
            Microsoft.Office.Interop.Excel.Range y4Range = xlApp.get_Range("E4:E784");
            Microsoft.Office.Interop.Excel.Range y5Range = xlApp.get_Range("F4:F784");
            Microsoft.Office.Interop.Excel.Range y6Range = xlApp.get_Range("G4:G784");
            Microsoft.Office.Interop.Excel.Range y7Range = xlApp.get_Range("H4:H784");
            Microsoft.Office.Interop.Excel.Range y8Range = xlApp.get_Range("I4:I784");
            Microsoft.Office.Interop.Excel.Range y9Range = xlApp.get_Range("J4:J784");
            Microsoft.Office.Interop.Excel.Range y10Range = xlApp.get_Range("K4:K784");
            int rowCount1 = xRange.Rows.Count;
            int colCount1 = xRange.Columns.Count;



            for (int i = 4; i <= rowCount1; i++)
            {
                string px = System.Convert.ToString(xRange.Cells[i, 1].Value2.ToString());
                xList1.Add(px);
            }
            for (int i = 4; i <= rowCount1; i++)
            {
                double py1 = System.Convert.ToDouble(y1Range.Cells[i, 1].Value2.ToString());
                double py2 = System.Convert.ToDouble(y2Range.Cells[i, 1].Value2.ToString());
                double py3 = System.Convert.ToDouble(y3Range.Cells[i, 1].Value2.ToString());
                double py4 = System.Convert.ToDouble(y4Range.Cells[i, 1].Value2.ToString());
                double py5 = System.Convert.ToDouble(y5Range.Cells[i, 1].Value2.ToString());
                double py6 = System.Convert.ToDouble(y6Range.Cells[i, 1].Value2.ToString());
                double py7 = System.Convert.ToDouble(y7Range.Cells[i, 1].Value2.ToString());
                double py8 = System.Convert.ToDouble(y8Range.Cells[i, 1].Value2.ToString());
                double py9 = System.Convert.ToDouble(y9Range.Cells[i, 1].Value2.ToString());
                double py10 = System.Convert.ToDouble(y10Range.Cells[i, 1].Value2.ToString());
                yList1.Add(py1);
                yList2.Add(py2);
                yList3.Add(py3);
                yList4.Add(py4);
                yList5.Add(py5);
                yList6.Add(py6);
                yList7.Add(py7);
                yList8.Add(py8);
                yList9.Add(py9);
                yList10.Add(py10);
            }

            chart1.Series[0].Points.DataBindXY(xList1, yList1);
            chart1.Series[1].Points.DataBindXY(xList1, yList2);
            chart1.Series[2].Points.DataBindXY(xList1, yList3);
            chart1.Series[3].Points.DataBindXY(xList1, yList4);
            chart1.Series[4].Points.DataBindXY(xList1, yList5);
            chart1.Series[5].Points.DataBindXY(xList1, yList6);
            chart1.Series[6].Points.DataBindXY(xList1, yList7);
            chart1.Series[7].Points.DataBindXY(xList1, yList8);
            chart1.Series[8].Points.DataBindXY(xList1, yList9);
            chart1.Series[9].Points.DataBindXY(xList1, yList10);
            this.chart1.Series[0].Color = Color.Yellow;
            this.chart1.Series[1].Color = Color.Red;
            this.chart1.Series[2].Color = Color.Aqua;
            this.chart1.Series[3].Color = Color.Black;
            this.chart1.Series[4].Color = Color.BurlyWood;
            this.chart1.Series[5].Color = Color.Chocolate;
            this.chart1.Series[6].Color = Color.DarkCyan;
            this.chart1.Series[7].Color = Color.DarkGreen;
            this.chart1.Series[8].Color = Color.DimGray;
            this.chart1.Series[9].Color = Color.HotPink;

            }
            catch (Exception)
            {
            }

        }//绘图

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }//confirm

    }
}
