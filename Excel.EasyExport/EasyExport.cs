using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Excel.EasyExport
{
    public static class EasyExport
    {
        public static ISheet CreateSheet<T>(IWorkbook workbook, List<ExportColumn> exportColums, List<T> data)
        {
            return CreateSheet(workbook, exportColums, data, "");
        }

        public static ISheet CreateSheet<T>(IWorkbook workbook, List<ExportColumn> exportColums, List<T> data, Dictionary<string, Func<T, int, object>> formatDic)
        {
            return CreateSheet(workbook, exportColums, data, "", formatDic);
        }

        public static ISheet CreateSheet<T>(IWorkbook workbook, List<ExportColumn> exportColums, List<T> data, ExportOptions exportOptions)
        {
            return CreateSheet(workbook, exportColums, data, "", null, exportOptions);
        }

        public static ISheet CreateSheet<T>(IWorkbook workbook, List<ExportColumn> exportColums, List<T> data, string sheetName, Dictionary<string, Func<T, int, object>> formatDic = null, ExportOptions exportOptions = null)
        {
            if (workbook == null)
                return null;

            if (exportColums == null || !exportColums.Any())
                return null;

            if (exportOptions == null)
                exportOptions = new ExportOptions();

            ISheet sheet = null;
            if (string.IsNullOrEmpty(sheetName))
                sheet = workbook.CreateSheet();
            else
                sheet = workbook.GetSheet(sheetName) ?? workbook.CreateSheet(sheetName);


            var rows = new List<IRow>();
            var rowPairs = new List<CellRangeAddress>();
            var colPairs = new List<CellRangeAddress>();

            int maxLevel = CreateHeader(workbook, sheet, exportColums, exportOptions, rowPairs, colPairs, rows);

            foreach (var item in rowPairs)
                sheet.AddMergedRegion(item);

            foreach (var item in colPairs)
                sheet.AddMergedRegion(item);

            foreach (var item in rows)
            {
                if (exportOptions != null)
                    item.Height = exportOptions.HeaderHeight;
            }

            if (data != null && data.Any())
                FillData(workbook, exportColums, sheet, data, maxLevel++, formatDic, exportOptions);

            SetColWidth(workbook, sheet, exportColums, data == null ? 0 : data.Count, exportOptions);

            return sheet;
        }


        private static int CreateHeader(IWorkbook workbook, ISheet sheet, List<ExportColumn> exportColums, ExportOptions exportOptions, List<CellRangeAddress> rowPairs, List<CellRangeAddress> colPairs, List<IRow> rows)
        {
            rows = new List<IRow>();
            rowPairs = rowPairs ?? new List<CellRangeAddress>();
            colPairs = colPairs ?? new List<CellRangeAddress>();

            ICellStyle cellStyle = workbook.CreateCellStyle();
            if (exportOptions.HeaderStyle != null)
                cellStyle = exportOptions.HeaderStyle;
            else
                cellStyle = GetDefaultHeaderStyle(workbook);

            //获取层级
            var columsLevels = GetColumLevel(exportColums);

            int maxLevel = columsLevels.Keys.Count;

            for (int i = 0; i < maxLevel; i++)
                rows.Add(sheet.CreateRow(i));

            int beginColumnIndex = 0;

            Dictionary<int, int> colPosDic = new Dictionary<int, int>();

            //循环顶级节点
            foreach (var item in columsLevels[0])
            {
                Dictionary<int, List<ExportColumn>> tempLevel = GetColumLevel(new List<ExportColumn>() { item });

                for (int i = 0; i < maxLevel; i++)//逐行
                {
                    if (!tempLevel.ContainsKey(i))
                        continue;

                    if (!colPosDic.ContainsKey(i))
                        colPosDic[i] = 0;

                    for (int j = 0; j < tempLevel[i].Count; j++)//逐个
                    {
                        var currentItem = tempLevel[i][j];

                        ICell cell = rows[i].CreateCell(colPosDic[i]);

                        //横向填充
                        if (tempLevel.ContainsKey(i))
                        {
                            string value = tempLevel[i].Count > j - beginColumnIndex ? tempLevel[i][j - beginColumnIndex].Label : "-";
                            cell.SetCellValue(value);
                            cell.CellStyle = cellStyle;

                            colPosDic[i]++;

                            var subLeaf = GetLeaf(currentItem);

                            if (subLeaf != null && subLeaf.Count > 1)
                            {
                                for (int y = 0; y < subLeaf.Count - 1; y++)
                                {
                                    ICell fillCell = rows[i].CreateCell(colPosDic[i]++);
                                    fillCell.SetCellValue("-");
                                }
                            }

                            if (subLeaf.Count > 1 && colPosDic[i] - subLeaf.Count != colPosDic[i] - 1)
                                rowPairs.Add(new CellRangeAddress(i, i, colPosDic[i] - subLeaf.Count, colPosDic[i] - 1));
                        }

                        //竖向填充
                        if (!tempLevel.ContainsKey(i) || currentItem.Children == null)
                        {

                            if (currentItem.Children == null && currentItem.Level < maxLevel - 1)
                            {
                                for (int y = 1; y < maxLevel - currentItem.Level; y++)
                                {
                                    if (!colPosDic.ContainsKey(i + y))
                                        colPosDic[i + y] = 0;

                                    ICell fillCell = rows[i + y].CreateCell(colPosDic[i] - 1);
                                    fillCell.SetCellValue("+");

                                    colPosDic[i + y]++;
                                }

                                colPairs.Add(new CellRangeAddress(i, maxLevel - 1, colPosDic[i] - 1, colPosDic[i] - 1));
                            }

                        }
                    }

                }
            }

            return maxLevel;

        }

        private static void SetColWidth(IWorkbook workbook, ISheet sheet, List<ExportColumn> exportColumns, long dataLength, ExportOptions exportOptions = null)
        {
            int next = 0;
            foreach (var column in exportColumns)
            {
                var leaves = GetLeaf(column);

                leaves.ForEach(item =>
                {
                    if (item.Width > 0)
                        sheet.SetColumnWidth(next, item.Width * 20);
                    else
                    {
                        if (exportOptions != null && exportOptions.AutoColSize && dataLength <= 5000)
                        {
                            sheet.AutoSizeColumn(next);
                            int currentWidth = sheet.GetColumnWidth(next);
                            sheet.SetColumnWidth(next, currentWidth + 300);
                        }

                    }

                    next++;
                });
            }
        }

        private static Dictionary<int, List<ExportColumn>> GetColumLevel(List<ExportColumn> exportColums, int level = 0, Dictionary<int, List<ExportColumn>> columLevels = null)
        {
            if (columLevels == null)
                columLevels = new Dictionary<int, List<ExportColumn>>();

            for (int i = 0; i < exportColums.Count; i++)
            {
                if (!columLevels.ContainsKey(level))
                    columLevels[level] = new List<ExportColumn>();

                var item = exportColums[i];
                item.Level = level;

                columLevels[level].Add(item);

                if (exportColums[i].Children != null && exportColums[i].Children.Any())
                    GetColumLevel(exportColums[i].Children, level + 1, columLevels);
            }

            return columLevels;
        }

        private static List<ExportColumn> GetLeaf(ExportColumn exportColumn)
        {
            if (exportColumn == null)
                return new List<ExportColumn>();


            List<ExportColumn> exportColumns = new List<ExportColumn>();

            if (exportColumn.Children == null || !exportColumn.Children.Any())
            {
                exportColumns.Add(exportColumn);
                return exportColumns;
            }

            for (int i = 0; i < exportColumn.Children.Count; i++)
                exportColumns.AddRange(GetLeaf(exportColumn.Children[i]));

            return exportColumns;
        }

        public static List<ExportColumn> GetAllLeaf(List<ExportColumn> columns)
        {
            List<ExportColumn> exportColumns = new List<ExportColumn>();


            foreach (var item in columns)
            {
                exportColumns.AddRange(GetLeaf(item));
            }

            return exportColumns;
        }

        private static void FillData<T>(IWorkbook wb, List<ExportColumn> exportColums, ISheet sheet, List<T> data, int nextRow, Dictionary<string, Func<T, int, object>> formatDic = null, ExportOptions exportOptions = null)
        {
            List<ExportColumn> allLeaves = new List<ExportColumn>();

            foreach (var item in exportColums)
                allLeaves.AddRange(GetLeaf(item));

            if (!allLeaves.Any())
                return;


            ICellStyle deafultDateStyle = GetDefaultDateTimeStyle(wb);

            Dictionary<int, ICellStyle> styleDic = new Dictionary<int, ICellStyle>();
            for (int rowIndex = 0; rowIndex < data.Count; rowIndex++)
            {
                var item = data[rowIndex];
                IRow row = sheet.CreateRow(nextRow++);
                for (int colIndex = 0; colIndex < allLeaves.Count; colIndex++)
                {
                    var leaf = allLeaves[colIndex];
                    ICell cell = row.CreateCell(colIndex);

                    object value = null;

                    //Format
                    if (formatDic != null && formatDic.ContainsKey(leaf.Prop))
                        value = formatDic[leaf.Prop](item, rowIndex);
                    else if (leaf.Format != null)
                        value = leaf.Format(item, rowIndex);
                    else
                        value = item.GetPropertyValue(allLeaves[colIndex].Prop);


                    FillCell(cell, value, deafultDateStyle);

                    if (leaf.Style != null && !styleDic.ContainsKey(colIndex))
                        styleDic.Add(colIndex, leaf.Style());

                    //全局数据样式
                    if (exportOptions != null && exportOptions.DataStyle != null && !styleDic.ContainsKey(colIndex))
                        cell.CellStyle = exportOptions.DataStyle;

                    //列样式优先于全局数据样式
                    if (styleDic.ContainsKey(colIndex))
                        cell.CellStyle = styleDic[colIndex];

                    if (exportOptions != null && exportOptions.CellCreateAfter != null)
                        exportOptions.CellCreateAfter(cell, colIndex, row, rowIndex);

                }

                if (exportOptions != null)
                {
                    row.Height = exportOptions.DataRowHeight;

                    if (exportOptions.RowCreateAfter != null)
                        exportOptions.RowCreateAfter(row, rowIndex);
                }
            }
        }

        private static void FillCell(ICell cell, object value, ICellStyle defaultDateStyle)
        {
            if (value == null)
                cell.SetCellValue((string)null);
            else if (value is DateTime)
            {
                cell.SetCellValue((DateTime)value);
                cell.CellStyle = defaultDateStyle;
            }
            else if (value.GetType().IsNumeric())
                cell.SetCellValue(Convert.ToDouble(value));
            else if (value is bool)
                cell.SetCellValue((bool)value);
            else
                cell.SetCellValue(value.ToString());
        }

        public static ICellStyle GetDefaultDateTimeStyle(IWorkbook wb)
        {
            IDataFormat dataformat = wb.CreateDataFormat();

            ICellStyle dateStyle = wb.CreateCellStyle();
            dateStyle.VerticalAlignment = VerticalAlignment.Center;
            dateStyle.DataFormat = dataformat.GetFormat("yyyy-MM-dd HH:mm:ss");

            return dateStyle;
        }

        public static ICellStyle GetDefaultHeaderStyle(IWorkbook wb)
        {
            var cellStyle = wb.CreateCellStyle();
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            IFont font = wb.CreateFont();
            font.IsBold = true;
            cellStyle.SetFont(font);
            return cellStyle;
        }

    }


}
