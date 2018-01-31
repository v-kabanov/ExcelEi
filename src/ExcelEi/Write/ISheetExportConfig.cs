// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-18
// Comment		
// **********************************************************************************************/

using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelEi.Write
{
    /// <summary>
    ///     Interface of an objects which defines how to export a data table into excel worksheet.
    /// </summary>
    public interface ISheetExportConfig
    {
        string SheetName { get; }

        string DataTableName { get; }

        /// <summary>
        ///     1 - based index of the first row to receive data.
        /// </summary>
        int TopSheetRowIndex { get; }

        /// <summary>
        ///     1 - based index of the leftmost column to receive data
        /// </summary>
        int LeftSheetColumnIndex { get; }

        bool ColumnHeaders { get; }

        Color? HeaderBackgroundColor { get; }

        BorderStyle? HeaderCellBorderStyle { get; }

        Color? HeaderCellBorderColor { get; }

        /// <summary>
        ///     Takes precedence over columns' vertical alignment setting.
        /// </summary>
        VerticalAlignment? HeaderVerticalAlignment { get; }

        /// <summary>
        ///     Whether to freeze rows/columns to the left and above the cell identified by the relative
        ///     0-based column index in the header. Null for not freezing.
        /// </summary>
        int? FreezeColumnIndex { get; }

        ICollection<IColumnExportConfig> Columns { get; }

        /// <summary>
        ///     Applied to whole sheet before exporting data.
        /// </summary>
        BorderStyle? DefaultCellBorderStyle { get; }

        /// <summary>
        ///     Applied to whole sheet before exporting data.
        /// </summary>
        Color? DefaultCellBorderColor { get; }

        /// <summary>
        ///     Sheet wide; see "Gridlines | View" setting in Excel 2013 "Page Layout" ribbon.
        /// </summary>
        bool? ShowGridlines { get; }

        /// <summary>
        ///     Sheet wide; see "Headings | View" setting in Excel 2013 "Page Layout" ribbon.
        /// </summary>
        bool? ShowHeadings { get; }

        /// <summary>
        ///     Extractor from row instance. Accepts data item and 0-based data row number.
        ///     Applied to all cells between first and last populated column in the row inclusive.
        /// </summary>
        /// <remarks>
        ///     Optional, may be null.
        /// </remarks>
        Func<object, int, Color?> DataRowCellBackgroundColorExtractor { get; }

        /// <summary>
        ///     Applied to all cells between first and last populated column in the row inclusive.
        /// </summary>
        BorderStyle? DataRowCellBorderStyle { get; }

        /// <summary>
        ///     Applied to all cells between first and last populated column in the row inclusive.
        /// </summary>
        Color? DataRowCellBorderColor { get; }
    }
}