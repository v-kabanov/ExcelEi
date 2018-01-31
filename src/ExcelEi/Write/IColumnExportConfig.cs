// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-18
// Comment		
// **********************************************************************************************/

using System;
using System.Drawing;

namespace ExcelEi.Write
{
    /// <summary>
    ///     Interface of an object which defines how to populate a worksheet column.
    /// </summary>
    public interface IColumnExportConfig
    {
        /// <summary>
        ///     Reference to containing sheet. 
        /// </summary>
        ISheetExportConfig SheetExportConfig { get; }

        /// <summary>
        ///     0 - based index of the sheet column, relative to the table position in the worksheet.
        /// </summary>
        int Index { get; }

        /// <summary>
        ///     Column caption.
        /// </summary>
        string Caption { get; }

        HorizontalAlignment? HorizontalAlignment { get; }

        /// <summary>
        ///     Note that <see cref="Write.VerticalAlignment.Justify"/> can be used to wrap text and auto-fit effectively, coupled with max width.
        /// </summary>
        VerticalAlignment? VerticalAlignment { get; }

        /// <summary>
        ///     Be careful when auto fitting AND wrapping. EPPlus sets width to the minumum allowed when wrapping text.
        /// </summary>
        bool? WrapText { get; }

        /// <summary>
        ///     Optional
        /// </summary>
        string Format { get; }

        BorderStyle? CellBorderStyle { get; }

        Color? CellBorderColor { get; }

        bool AutoFit { get; }

        /// <summary>
        ///     Minimum width (in characters) will be enforced regardless of <see cref="AutoFit"/>.
        /// </summary>
        double? MinimumWidth { get; }

        /// <summary>
        ///      In characters, will be honored when <see cref="AutoFit"/> is true.
        /// </summary>
        double? MaximumWidth { get; }

        /// <summary>
        ///     Extractor from row instance.
        /// </summary>
        /// <remarks>
        ///     Mandatory, must never be null.
        /// </remarks>
        Func<object, object> ValueExtractor { get; }

        /// <summary>
        ///     Extractor from row instance.
        /// </summary>
        /// <remarks>
        ///     Optional, may be null.
        /// </remarks>
        Func<object, Color?> FontColorExtractor { get; }

        /// <summary>
        ///     Extractor from row instance. Accepts data item and row number.
        /// </summary>
        /// <remarks>
        ///     Optional, may be null.
        /// </remarks>
        Func<object, int, Color?> BackgroundColorExtractor { get; }

        /// <summary>
        ///     Extractor from row instance.
        /// </summary>
        /// <remarks>
        ///     Optional, may be null.
        /// </remarks>
        Func<object, string> CellCommentExtractor { get; }
    }
}