using AnyCAD.Platform;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Interface
{
    class ViewParametrs
    {
        /// <summary>
        /// 节点ID列表
        /// </summary>
        public static List<ElementId> IDs = new List<ElementId>();
        /// <summary>
        /// 当前节点ID
        /// </summary>
        public static ElementId CurrentId = new ElementId(0);
    }
}
