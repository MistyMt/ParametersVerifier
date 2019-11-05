using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Interface
{
    class Sensor
    {
        /// <summary>
        /// 名称
        /// </summary>
        public string name = "未命名";
        
        /// <summary>
        /// 编号
        /// </summary>
        public int serialNumber = 000000;
        
        /// <summary>
        /// 测量范围左
        /// </summary>
        public int rangeMin = 0000;
        
        /// <summary>
        /// 测量范围右
        /// </summary>
        public int rangeMax = 0000;
        
        /// <summary>
        /// 型号规格
        /// </summary>
        public string type = "未指定类型";
        
        /// <summary>
        /// 不确定度或准确度等级或最大允许误差
        /// </summary>
        public double uncertainty = 0.0;
        
        /// <summary>
        /// 证书编号/有效期
        /// </summary>
        public string certificateNo = "未指定编号";

    }
}
