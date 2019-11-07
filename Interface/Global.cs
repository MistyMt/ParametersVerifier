using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;


namespace Interface
{
    public static class Global
    {
        /// <summary>
        /// 数据分析面板
        /// </summary>
        public static List<FormDataView> dataViewForms = new List<FormDataView>();
       
        /// <summary>
        /// 基础信息面板
        /// </summary>
        public static List<FormBasicInformation> basicInformationForms = new List<FormBasicInformation>();

        /// <summary>
        /// 数据总览面板
        /// </summary>
        public static List<FormAnalysisOfData> analysisOfDataForms = new List<FormAnalysisOfData>();

        /// <summary>
        /// 对象名
        /// </summary>
        public static string objectName = "未选择";


        /// <summary>
        /// 模板文件名
        /// </summary>
        public static string templateName = "冷库验证项目模拟报告模板";

        /// <summary>
        /// 截图
        /// </summary>
        /// <returns></returns>
        public static Bitmap GetScreenCapture()
        {
            System.Drawing.Rectangle tScreenRect = new System.Drawing.Rectangle(0, 0, Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Bitmap tSrcBmp = new Bitmap(tScreenRect.Width, tScreenRect.Height); // 用于屏幕原始图片保存
            Graphics gp = Graphics.FromImage(tSrcBmp);
            gp.CopyFromScreen(0, 0, 0, 0, tScreenRect.Size);
            gp.DrawImage(tSrcBmp, 0, 0, tScreenRect, GraphicsUnit.Pixel);
            return tSrcBmp;
        }

        public static List<Sensor> sensors = new List<Sensor>();




    }
}
