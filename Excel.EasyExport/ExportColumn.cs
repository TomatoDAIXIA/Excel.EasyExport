using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using Excel.EasyExport;

namespace Excel.EasyExport
{
    public class ExportColumn
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// 字段
        /// </summary>
        public string Prop { get; set; }

        /// <summary>
        /// 列宽
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// 自定义Style
        /// </summary>
        public Func<ICellStyle> Style { get; set; }

        /// <summary>
        /// 自定义格式化，
        /// 第一个参数是当前行数据
        /// 第二个参数是当前行索引
        /// 第三个参数是返回值
        /// </summary>
        public Func<object, int, object> Format { get; set; }

        /// <summary>
        /// 当前表头层级
        /// </summary>
        internal int Level { get; set; }

        public List<ExportColumn> Children { get; set; }
    }

}
