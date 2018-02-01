// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-19
// Comment		
// **********************************************************************************************/

using System;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using ExcelEi.Read;

namespace ExcelEi.Write
{
    /// <summary>
    ///     Configuration for exporting 1 column out of data source such as <see cref="DataTable"/> or POCO collection.
    /// </summary>
    public class DataColumnExportAutoConfig : IColumnExportConfig
    {
        public const string DefaultDateTimeFormat = "dd-MMM-yyyy hh:mm:ss AM/PM";

        /// <param name="sheet">
        ///     Mandatory, containing sheet export configuration
        /// </param>
        /// <param name="column">
        ///     Source data column
        /// </param>
        public DataColumnExportAutoConfig(ISheetExportConfig sheet, DataColumn column)
            : this(sheet, column, column.Ordinal)
        {
        }

        /// <summary>
        ///     Initializes new instance from existing <see cref="DataColumn"/> and explicitly specified column position in the sheet.
        /// </summary>
        /// <param name="sheet">
        ///     Mandatory, containing sheet export configuration
        /// </param>
        /// <param name="column">
        ///     Source data column
        /// </param>
        /// <param name="sheetColumnIndex">
        ///     0-based relative index of the column in the sheet
        /// </param>
        public DataColumnExportAutoConfig(ISheetExportConfig sheet, DataColumn column, int sheetColumnIndex)
            : this(sheet, column, sheetColumnIndex, column.ColumnName)
        {
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="DataColumnExportAutoConfig"/> class.
        /// </summary>
        /// <param name="sheet">
        ///     Mandatory
        /// </param>
        /// <param name="column">
        ///     ADO.NET data column from which the sheet column is going to be populated.
        /// </param>
        /// <param name="sheetColumnIndex">
        ///     0-based relative sheet column index to receive data.
        /// </param>
        public DataColumnExportAutoConfig(ISheetExportConfig sheet, DataColumn column, int sheetColumnIndex, string caption)
            : this(sheet, sheetColumnIndex, caption, new DataColumnSource(column))
        {
        }

        /// <summary>
        ///     Lower-level initializer which can work with any data item type.
        /// </summary>
        /// <param name="sheet">
        ///     Mandatory
        /// </param>
        /// <param name="index">
        ///     Relative 0-based sheet column index
        /// </param>
        /// <param name="caption">
        ///     Column header text
        /// </param>
        /// <param name="columnDataSource">
        ///     Mandatory
        /// </param>
        public DataColumnExportAutoConfig(ISheetExportConfig sheet, int index, string caption, IColumnDataSource columnDataSource)
        {
            Check.DoRequireArgumentNotNull(sheet, nameof(sheet));
            Check.DoRequireArgumentNotNull(caption, nameof(caption));
            Check.DoRequireArgumentNotNull(columnDataSource, nameof(columnDataSource));

            ColumnDataSource = columnDataSource;
            SheetExportConfig = sheet;

            Index = index;
            Caption = caption;

            var isPrimitive = IsPrimitive(columnDataSource.DataType);
            var isNumeric = isPrimitive && IsNumeric(columnDataSource.DataType);
            var isDateTime = !isNumeric && IsDateTime(columnDataSource.DataType);

            HorizontalAlignment = (isDateTime || isPrimitive) ? Write.HorizontalAlignment.Right : Write.HorizontalAlignment.Left;
            VerticalAlignment = Write.VerticalAlignment.Top;

            AutoFit = true;

            if (isDateTime)
            {
                Format = DefaultDateTimeFormat;
            }
            else if (columnDataSource.DataType == typeof(string))
            {
                MaximumWidth = 50;
            }
        }

        public IColumnDataSource ColumnDataSource { get; }

        /// <summary>
        ///     Reference to containing sheet. 
        /// </summary>
        public ISheetExportConfig SheetExportConfig { get; set; }

        /// <summary>
        ///     0 - based index of the column, relative to the table position in the worksheet.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        ///     Column caption.
        ///     Not used by default value extractor, therefore can be safely changed.
        /// </summary>
        public string Caption { get; set; }

        public HorizontalAlignment? HorizontalAlignment { get; set; }

        public VerticalAlignment? VerticalAlignment { get; set; }

        public bool? WrapText { get; set; }

        /// <summary>
        ///     Optional
        /// </summary>
        public string Format { get; set; }

        public BorderStyle? CellBorderStyle { get; set; }

        public Color? CellBorderColor { get; set; }

        public bool AutoFit { get; set; }

        /// <summary>
        ///     Minimum width in characters will be enforced regardless of <see cref="IColumnExportConfig.AutoFit"/>.
        /// </summary>
        public double? MinimumWidth { get; set; }

        /// <summary>
        ///     Width in characters, will be honored when <see cref="IColumnExportConfig.AutoFit"/> is true.
        /// </summary>
        public double? MaximumWidth { get; set; }

        /// <inheritdoc />
        public object GetCellValue(object dataObject) => ColumnDataSource.GetValue(dataObject);

        /// <summary>
        ///     Extractor from row instance.
        /// </summary>
        /// <remarks>
        ///     Optional, may be null.
        /// </remarks>
        public Func<object, Color?> FontColorExtractor { get; set; }

        /// <summary>
        ///     Extractor from row instance.
        /// </summary>
        /// <remarks>
        ///     Optional, may be null.
        /// </remarks>
        public Func<object, int, Color?> BackgroundColorExtractor { get; set; }

        /// <summary>
        ///     Extractor from row instance.
        /// </summary>
        /// <remarks>
        ///     Optional, may be null.
        /// </remarks>
        public Func<object, string> CellCommentExtractor { get; set; }

        private bool IsDateTime(Type type)
        {
            type = Nullable.GetUnderlyingType(type) ?? type;

            return typeof(SqlDateTime).IsAssignableFrom(type) || typeof(DateTime).IsAssignableFrom(type);
        }

        private bool IsPrimitive(Type type)
        {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return type.IsPrimitive;
        }

        private bool IsNumeric(Type type)
        {
            type = Nullable.GetUnderlyingType(type) ?? type;

            return type.IsPrimitive
                && type != typeof(bool)
                && type != typeof(char)
                && type != typeof(IntPtr)
                && type != typeof(UIntPtr);
        }
    }
}