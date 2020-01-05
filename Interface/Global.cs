using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;
using AnyCAD.Platform;


namespace Interface
{
    public static class Global
    {
        #region 面板
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
        public static List<Form> analysisOfDataForms = new List<Form>();
        #endregion

        /// <summary>
        /// 对象名
        /// </summary>
        public static string objectName = string.Empty;


        /// <summary>
        /// 模板文件名
        /// </summary>
        public static string templateName = string.Empty;

        /// <summary>
        /// excel数据源名
        /// </summary>
        public static string dataSourceName = string.Empty;

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

        /// <summary>
        /// 测试布点列表
        /// string:node的name属性
        /// Sensor:测点传感器对象
        /// </summary>
        public static Dictionary<string, Sensor> sensors = new Dictionary<string, Sensor>();


        /// <summary>
        /// 对象大小,单位m
        /// </summary>
        public static double objectLen = 0;
        public static double objectWidth = 0;
        public static double objectHeight = 0;

        /// <summary>
        /// 渲染窗口
        /// </summary>
        public static AnyCAD.Presentation.RenderWindow3d RenderView = null;

        /// <summary>
        /// 温度上界
        /// </summary>
        public static double tempLimitHigh = 0.0;

        /// <summary>
        /// 温度下界
        /// </summary>
        public static double tempLimitLow = 0.0;


    }
}
