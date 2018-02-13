// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-19
// Comment		
// **********************************************************************************************/

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;

namespace ExcelEi.Write
{
    /// <summary>
    ///     Builds configuration for a single table.
    /// </summary>
    public class DataTableExportAutoConfig : ISheetExportConfig
    {
        private static readonly Color? DefaultEvenBackgroundColor = Color.White;
        private static readonly Color? DefaultOddBackgroundColor = ColorTranslator.FromHtml("#FAF8F8");

        private readonly List<IColumnExportConfig> _columns = new List<IColumnExportConfig>();

#if !NOADONET

        /// <summary>
        ///     Creates default configuration from existing ADO.NET <see cref="DataTable"/>.
        /// </summary>
        public DataTableExportAutoConfig(DataTable dataTable)
            : this()
        {
            SheetName = dataTable.TableName;
            DataTableName = dataTable.TableName;

            foreach (DataColumn column in dataTable.Columns)
            {
                AddColumn(new DataColumnExportAutoConfig(this, column));
            }
        }
#endif

        /// <summary>
        ///     Initialize new instance with defaults and empty columns collection.
        /// </summary>
        public DataTableExportAutoConfig()
        {
            EvenRowBackgroundColor = DefaultEvenBackgroundColor;
            OddRowBackgroundColor = DefaultOddBackgroundColor;

            HeaderBackgroundColor = DefaultOddBackgroundColor;
            HeaderCellBorderStyle = BorderStyle.Thin;
            HeaderCellBorderColor = Color.LightGray;

            TopSheetRowIndex = 1;
            LeftSheetColumnIndex = 1;
            ColumnHeaders = true;
            DataRowCellBackgroundColorExtractor = (o, n) => (n % 2 == 0) ? EvenRowBackgroundColor : OddRowBackgroundColor;

            DataRowCellBorderStyle = BorderStyle.Thin;
            DataRowCellBorderColor = Color.LightGray;

            FreezeColumnIndex = 1;

            HeaderVerticalAlignment = VerticalAlignment.Top;
        }

        /// <summary>
        ///     Cancel default formatting for header.
        /// </summary>
        public void DisableHeaderBorderAndBackgroundFormatting()
        {
            HeaderBackgroundColor = null;
            HeaderCellBorderStyle = null;
            HeaderCellBorderColor = null;
        }

        /// <summary>
        ///     EPPlus is slow when formatting. This is a convenience method to disable expensive formatting
        ///     when data set is big and performance important.
        /// </summary>
        public void DisableDataRowCellBorderAndBackgroundFormatting()
        {
            DataRowCellBackgroundColorExtractor = null;
            DataRowCellBorderStyle = null;
            DataRowCellBorderColor = null;
        }

        /// <summary>
        ///     EPPlus is slow when formatting. This is a convenience method to disable expensive formatting
        ///     when data set is big and performance important.
        /// </summary>
        public void DisableCellBorderAndBackgroundFormatting()
        {
            DisableHeaderBorderAndBackgroundFormatting();
            DisableDataRowCellBorderAndBackgroundFormatting();
        }

        public DataTableExportAutoConfig AddColumn(IColumnExportConfig columnConfig)
        {
            Check.DoRequireArgumentNotNull(columnConfig, "columnConfig");
            Check.DoCheckArgument(GetColumnBySheetIndex(columnConfig.Index) == null, "The sheet column population is already configured");

            _columns.Add(columnConfig);
            return this;
        }


        public IColumnExportConfig GetColumnBySheetIndex(int index)
        {
            return _columns.FirstOrDefault(c => c.Index == index);
        }

        /// <summary>
        ///     Get column config by caption for further customisation.
        /// </summary>
        /// <param name="caption">
        ///     Case sensitive, to be compared with <see cref="IColumnExportConfig.Caption"/>.
        /// </param>
        /// <returns>
        ///     Config or null.
        /// </returns>
        public DataColumnExportAutoConfig GetAutoColumnConfig(string caption)
        {
            return (DataColumnExportAutoConfig)Columns.FirstOrDefault(c => c.Caption == caption);
        }

        public IList<IColumnExportConfig> Columns => _columns;

        public string SheetName { get; set; }

        public string DataTableName { get; set; }

        /// <summary>
        ///     1 - based index of the first row to receive data.
        /// </summary>
        public int TopSheetRowIndex { get; set; }

        /// <summary>
        ///     1 - based index of the leftmost column to receive data
        /// </summary>
        public int LeftSheetColumnIndex { get; set; }

        /// <summary>
        ///     Takes precedence over columns' vertical alignment setting.
        /// </summary>
        public VerticalAlignment? HeaderVerticalAlignment { get; set; }

        public bool ColumnHeaders { get; set; }

        public Color? HeaderBackgroundColor { get; set; }

        /// <inheritdoc />
        public BorderStyle? HeaderCellBorderStyle { get; set; }

        /// <inheritdoc />
        public Color? HeaderCellBorderColor { get; set; }

        /// <summary>
        ///     Whether to freeze rows/columns to the left and above the cell identified by the relative
        ///     0-based column index in the header. Null for not freezing.
        /// </summary>
        public int? FreezeColumnIndex { get; set; }

        ICollection<IColumnExportConfig> ISheetExportConfig.Columns => _columns;

        /// <summary>
        ///     Sheet wide; see "Gridlines | View" setting in Excel 2013 "Page Layout" ribbon.
        /// </summary>
        public bool? ShowGridlines { get; set; }

        /// <summary>
        ///     Sheet wide; see "Headings | View" setting in Excel 2013 "Page Layout" ribbon.
        /// </summary>
        public bool? ShowHeadings { get; set; }

        public BorderStyle? DefaultCellBorderStyle { get; set; }

        public Color? DefaultCellBorderColor { get; set; }

        /// <summary>
        ///     Extractor from row instance. Accepts data item and 0-based data row number.
        ///     Applied to all cells between first and last populated column in the row inclusive.
        /// </summary>
        /// <remarks>
        ///     Optional, may be null.
        /// </remarks>
        public Func<object, int, Color?> DataRowCellBackgroundColorExtractor { get; set; }

        /// <summary>
        ///     Applied to all cells between first and last populated column in the row inclusive.
        /// </summary>
        public BorderStyle? DataRowCellBorderStyle { get; set; }

        /// <summary>
        ///     Applied to all cells between first and last populated column in the row inclusive.
        /// </summary>
        public Color? DataRowCellBorderColor { get; set; }

        public Color? EvenRowBackgroundColor { get; set; }

        public Color? OddRowBackgroundColor { get; set; }

    }
}