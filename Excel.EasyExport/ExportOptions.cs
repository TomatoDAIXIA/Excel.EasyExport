using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.EasyExport
{
    public class ExportOptions
    {
        public short HeaderHeight { get; set; } = 400;
        public short DataRowHeight { get; set; } = 300;
        public ICellStyle HeaderStyle { get; set; }

        public ICellStyle DataStyle { get; set; }

        public Action<IRow, int> RowCreateAfter { get; set; }

        public Action<ICell, int, IRow, int> CellCreateAfter { get; set; }

        public bool AutoColSize { get; set; } = true;

    }
}
