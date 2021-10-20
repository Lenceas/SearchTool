using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SearchTool
{
    /// <summary>
    /// Excel通用模型
    /// </summary>
    public class ExcelModel
    {
        /// <summary>
        /// 序号
        /// </summary>
        public string? id { get; set; }

        /// <summary>
        /// 类型
        /// </summary>
        public string type { get; set; }

        /// <summary>
        /// 集合
        /// </summary>
        public string item { get; set; }
    }
}
