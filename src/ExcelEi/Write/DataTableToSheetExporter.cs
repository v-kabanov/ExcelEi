// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-19
// Comment		
// **********************************************************************************************/

using System.Drawing;
using System.Linq;
using System.Reflection;
using ExcelEi.Read;
#if !DISABLE_LOG4NET
using log4net;
#endif
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelEi.Write
{
    /// <summary>
    ///     Implements export from <see cref="IDataTable"/> into excel worksheet.
    /// </summary>
    public class DataTableToSheetExporter
    {
        public const int MaxExcelSheetRowCount = ExcelPackage.MaxRows;

#if !DISABLE_LOG4NET
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
#endif

        private int _currentSheetRowIndex;

        private readonly int _minPopulatedSheetColumnIndex;
        private readonly int _maxPopulatedSheetRowIndex;

        /// <summary>
        ///     Initializes a new instance of the <see cref="DataTableToSheetExporter"/> class.
        /// </summary>
        public DataTableToSheetExporter(ExcelWorksheet excelWorksheet, ISheetExportConfig sheetExportConfig, IDataTable dataTable)
        {
            Check.DoRequireArgumentNotNull(excelWorksheet, "excelWorksheet");
            Check.DoRequireArgumentNotNull(sheetExportConfig, "sheetExportConfig");
            Check.DoRequireArgumentNotNull(dataTable, "dataTable");

            ExcelWorksheet = excelWorksheet;
            SheetExportConfig = sheetExportConfig;

            _minPopulatedSheetColumnIndex = GetSheetColumnIndex(SheetExportConfig.Columns.Min(c => c.Index));
            _maxPopulatedSheetRowIndex = GetSheetColumnIndex(SheetExportConfig.Columns.Max(c => c.Index));
            
            DataTable = dataTable;
        }

        public ExcelWorksheet ExcelWorksheet { get; }
        public ISheetExportConfig SheetExportConfig { get; }
        public IDataTable DataTable { get; }

        public bool IfSparseColumns()
        {
            var columnIndexes = SheetExportConfig.Columns.Select(c => c.Index).OrderBy(i => i).ToArray();
            for (var i = 0; i < columnIndexes.Length - 1; ++i)
                if (columnIndexes[i] + 1 != columnIndexes[i + 1])
                    return true;

            return false;
        }

        public void Export()
        {
#if !DISABLE_LOG4NET
            Log.DebugFormat("Exporting {0}", SheetExportConfig.SheetName);
#endif
            _currentSheetRowIndex = SheetExportConfig.TopSheetRowIndex;

            SetUpSheet();
            SetUpColumns(SheetExportConfig.ColumnHeaders);

            var dataRowNumber = 0;

            foreach (var dataItem in DataTable.Rows)
            {
                var row = GetContiguousCurrentSheetCellRow();

                var color = SheetExportConfig.DataRowCellBackgroundColorExtractor?.Invoke(dataItem, dataRowNumber);
                if (color.HasValue)
                {
                    var rowStyle = row.Style;
                    rowStyle.Fill.PatternType = ExcelFillStyle.Solid;
                    rowStyle.Fill.BackgroundColor.SetColor(color.Value);
                }

                SetUpBorder(row, SheetExportConfig.DataRowCellBorderColor, SheetExportConfig.DataRowCellBorderStyle);

                foreach (var columnConfig in SheetExportConfig.Columns)
                {
                    var sheetColumnIndex = GetSheetColumnIndex(columnConfig);

                    var cell = ExcelWorksheet.Cells[_currentSheetRowIndex, sheetColumnIndex];
                    ExcelStyle cellStyle = null;
                    var fontColor = columnConfig.FontColorExtractor?.Invoke(dataItem);
                    var backgroundColor = columnConfig.BackgroundColorExtractor?.Invoke(dataItem, dataRowNumber);
                    if (backgroundColor.HasValue
                        || fontColor.HasValue
                        || columnConfig.CellBorderColor.HasValue
                        || columnConfig.CellBorderStyle.HasValue)
                        cellStyle = cell.Style;

                    cell.Value = columnConfig.GetCellValue(dataItem);

                    if (columnConfig.CellCommentExtractor != null)
                    {
                        var comment = cell.AddComment(columnConfig.CellCommentExtractor.Invoke(dataItem), string.Empty);
                        comment.AutoFit = true;
                    }

                    FormatCell(cellStyle, backgroundColor, fontColor, columnConfig);
                }

                ++_currentSheetRowIndex;
                ++dataRowNumber;
                if (_currentSheetRowIndex >= (MaxExcelSheetRowCount - 5))
                {
#if !DISABLE_LOG4NET
                    Log.WarnFormat("Stopping export at row {0} so that not to exceed excel limit", dataRowNumber);
#endif
                    break;
                }
            }

            var lastDataRowIndex = _currentSheetRowIndex - 1;
            if (SheetExportConfig.DataRowCellBorderColor.HasValue || SheetExportConfig.DataRowCellBorderStyle.HasValue)
                foreach (var columnConfig in SheetExportConfig.Columns)
                {
                    var columnIndex = GetSheetColumnIndex(columnConfig);
                    var column = ExcelWorksheet.Cells[SheetExportConfig.TopSheetRowIndex, columnIndex, lastDataRowIndex, columnIndex];
                    SetUpBorder(column, SheetExportConfig.DataRowCellBorderColor, SheetExportConfig.DataRowCellBorderStyle);
                }

            FormatAfterDataExport();

#if !DISABLE_LOG4NET
            Log.DebugFormat("Finished");
#endif
        }

        private void FormatCell(ExcelStyle excelStyle, Color? backgroundColor, Color? fontColor, IColumnExportConfig columnConfig)
        {
            if (excelStyle == null)
                return;

            if (backgroundColor.HasValue)
            {
                excelStyle.Fill.PatternType = ExcelFillStyle.Solid;
                excelStyle.Fill.BackgroundColor.SetColor(backgroundColor.Value);
            }

            if (fontColor.HasValue)
            {
                excelStyle.Font.Color.SetColor(fontColor.Value);
            }

            SetUpBorder(excelStyle, columnConfig.CellBorderColor, columnConfig.CellBorderStyle);
        }

        private ExcelRange GetContiguousCurrentSheetCellRow()
        {
            return GetContiguousSheetCellRow(_currentSheetRowIndex);
        }

        private ExcelRange GetContiguousSheetCellRow(int sheetRowIndex)
        {
            return ExcelWorksheet.Cells[sheetRowIndex, _minPopulatedSheetColumnIndex, sheetRowIndex, _maxPopulatedSheetRowIndex];
        }

        private void FormatAfterDataExport()
        {
            // autofit columns
            foreach (var columnConfig in SheetExportConfig.Columns)
            {
                if (columnConfig.AutoFit)
                {
                    var sheetColumnIndex = GetSheetColumnIndex(columnConfig);
                    var excelColumn = ExcelWorksheet.Column(sheetColumnIndex);

                    if (columnConfig.MinimumWidth.HasValue || columnConfig.MaximumWidth.HasValue)
                    {
                        if (columnConfig.MaximumWidth.HasValue)
                        {
                            var minimumWidth = columnConfig.MinimumWidth ?? 0;
                            excelColumn.AutoFit(minimumWidth, columnConfig.MaximumWidth.Value);
                        }
                        else
                        {
                            excelColumn.AutoFit(columnConfig.MinimumWidth.Value);
                        }
                    }
                    else
                    {
                        excelColumn.AutoFit();
                    }
                }
            }
        }

        private void SetUpSheet()
        {
            SetUpBorder(ExcelWorksheet.Cells, SheetExportConfig.DefaultCellBorderColor, SheetExportConfig.DefaultCellBorderStyle);

            if (SheetExportConfig.ShowGridlines.HasValue)
            {
                ExcelWorksheet.View.ShowGridLines = SheetExportConfig.ShowGridlines.Value;
            }

            if (SheetExportConfig.ShowHeadings.HasValue)
            {
                ExcelWorksheet.View.ShowHeaders = SheetExportConfig.ShowHeadings.Value;
            }
        }

        private void SetUpBorder(ExcelRange range, Color? color, BorderStyle? style)
        {
            if (color == null && style == null)
                return;

            style = style ?? BorderStyle.Thin;
            if (style == BorderStyle.None)
            {
                range.Style.Border.BorderAround(ExcelBorderStyle.None);
                return;
            }

            if (color == null)
            {
                range.Style.Border.BorderAround((ExcelBorderStyle)((int)style.Value));
                return;
            }

            range.Style.Border.BorderAround((ExcelBorderStyle)((int)style.Value), color.Value);
        }

        private void SetUpBorder(ExcelStyle excelStyle, Color? color, BorderStyle? style)
        {
            if (color == null && style == null)
                return;

            style = style ?? BorderStyle.Thin;
            if (style == BorderStyle.None)
            {
                excelStyle.Border.BorderAround(ExcelBorderStyle.None);
                return;
            }

            if (color == null)
            {
                excelStyle.Border.BorderAround((ExcelBorderStyle)((int)style.Value));
                return;
            }

            excelStyle.Border.BorderAround((ExcelBorderStyle)((int)style.Value), color.Value);
        }

        private void SetUpColumns(bool createHeaderRow)
        {
            var headerRow = GetContiguousCurrentSheetCellRow();

            if (createHeaderRow)
            {
                foreach (var columnConfig in SheetExportConfig.Columns)
                {
                    var columnIndex = GetSheetColumnIndex(columnConfig.Index);
                    var cell = ExcelWorksheet.Cells[_currentSheetRowIndex, columnIndex];
                    SetUpBorder(cell, SheetExportConfig.HeaderCellBorderColor, SheetExportConfig.HeaderCellBorderStyle);
                }

                if (SheetExportConfig.FreezeColumnIndex.HasValue)
                {
                    ExcelWorksheet.View.FreezePanes(_currentSheetRowIndex + 1, SheetExportConfig.FreezeColumnIndex.Value + 1);
                }

                if (SheetExportConfig.HeaderBackgroundColor.HasValue)
                {
                    headerRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRow.Style.Fill.BackgroundColor.SetColor(SheetExportConfig.HeaderBackgroundColor.Value);
                }
            }

            foreach (var columnConfig in SheetExportConfig.Columns)
            {
                var sheetColumnIndex = GetSheetColumnIndex(columnConfig);

                if (createHeaderRow)
                {
                    var headerCell = ExcelWorksheet.Cells[_currentSheetRowIndex, sheetColumnIndex];

                    headerCell.Value = columnConfig.Caption;

                    SetUpBorder(headerCell, columnConfig.CellBorderColor, columnConfig.CellBorderStyle);
                }

                var excelColumn = ExcelWorksheet.Column(sheetColumnIndex);

                var excelColumnStyle = excelColumn.Style;

                if (columnConfig.HorizontalAlignment.HasValue)
                {
                    excelColumnStyle.HorizontalAlignment = (ExcelHorizontalAlignment)((int)columnConfig.HorizontalAlignment.Value);
                }

                if (columnConfig.VerticalAlignment.HasValue)
                {
                    excelColumnStyle.VerticalAlignment = (ExcelVerticalAlignment)((int)columnConfig.VerticalAlignment.Value);
                }

                if (columnConfig.WrapText.HasValue)
                {
                    excelColumnStyle.WrapText = columnConfig.WrapText.Value;
                }

                if (!string.IsNullOrEmpty(columnConfig.Format))
                {
                    excelColumnStyle.Numberformat.Format = columnConfig.Format;
                }

                if (columnConfig.MinimumWidth.HasValue)
                {
                    excelColumn.Width = columnConfig.MinimumWidth.Value;
                }
            }

            if (createHeaderRow)
            {
                //var headerRow = ExcelWorksheet.Row(_currentSheetRowIndex);
                headerRow.Style.Font.Bold = true;
                ++_currentSheetRowIndex;
                if (SheetExportConfig.HeaderVerticalAlignment.HasValue)
                {
                    headerRow.Style.VerticalAlignment = (ExcelVerticalAlignment)((int)SheetExportConfig.HeaderVerticalAlignment.Value);
                }
            }
        }

        /// <summary>
        ///     Get 1-based column index of the column in the sheet.
        /// </summary>
        private int GetSheetColumnIndex(IColumnExportConfig columnConfig)
        {
            return GetSheetColumnIndex(columnConfig.Index);
        }

        /// <summary>
        ///     Get 1-based column index of the column in the sheet.
        /// </summary>
        /// <param name="relativeDataColumnIndex">
        ///     Value from <see cref="IColumnExportConfig.Index"/>
        /// </param>
        /// <returns></returns>
        private int GetSheetColumnIndex(int relativeDataColumnIndex)
        {
            return SheetExportConfig.LeftSheetColumnIndex + relativeDataColumnIndex;
        }
    }
}