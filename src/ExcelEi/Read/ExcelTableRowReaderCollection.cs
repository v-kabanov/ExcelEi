// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-01-17
// Comment		
// **********************************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Represents range of rows from EPPlus excel worksheet as a collection of rows with named columns.
    ///     Virtual table in worksheet can have sparse columns by virtue of column index resolver.
    /// </summary>
    /// <remarks>
    ///     Allows and skips single empty rows, stops reading when 2 consecutive empty rows are encountered.
    ///     This is due to legacy formats with empty first row.
    /// </remarks>
    public class ExcelTableRowReaderCollection : ITableRowReaderCollection
    {
        /// <summary>
        ///     Max count as well as index since it's 1-based
        /// </summary>
        public const int MaxExcelRowCount = ExcelPackage.MaxRows;

        /// <summary>
        ///     Max count as well as index since it's 1-based
        /// </summary>
        public const int MaxExcelColumnCount = ExcelPackage.MaxColumns;

        public ExcelWorksheet ExcelWorksheet { get; }

        private readonly int _startRowIndex;
        private readonly int? _endRowIndex;
        /// <summary>
        ///     Index resolves name to <see cref="ExcelWorksheet"/> column index which can be passed to <see cref="ExcelWorksheet.Cells"/>
        /// </summary>
        private readonly IDictionary<string, int> _columnNameIndex;

        /// <param name="startRowIndex">
        ///     1-based index as understood by <see cref="ExcelWorksheet"/>, inclusive.
        /// </param>
        /// <param name="endRowIndex">
        ///     Optional, 1-based index as understood by <see cref="ExcelWorksheet"/>, exclusive to be able to reference empty tables.
        /// </param>
        /// <param name="excelWorksheet">
        ///     Worksheet where table is defined
        /// </param>
        /// <param name="columnNameIndex">
        ///     Mandatory, maps column names to indexec as understood by <see cref="ExcelWorksheet"/>.
        /// </param>
        public ExcelTableRowReaderCollection(int startRowIndex, int? endRowIndex, ExcelWorksheet excelWorksheet, IDictionary<string, int> columnNameIndex)
        {
            Check.DoRequireArgumentNotNull(excelWorksheet, nameof(excelWorksheet));
            Check.DoRequireArgumentNotNull(columnNameIndex, nameof(columnNameIndex));
            Check.DoCheckArgument(startRowIndex >= 1 , "Invalid row index, must be 1..1048576 inclusive");
            Check.DoCheckArgument(!endRowIndex.HasValue || (endRowIndex >= startRowIndex && endRowIndex <= MaxExcelRowCount + 1)
                , "Invalid row range");

            _startRowIndex = startRowIndex;
            _endRowIndex = endRowIndex;
            ExcelWorksheet = excelWorksheet;
            _columnNameIndex = columnNameIndex;

            if (_endRowIndex.HasValue)
                Count = _endRowIndex - _startRowIndex;
        }

        /// <summary>
        ///     Get or set max number of cells in the row which still allow to consider it empty.
        /// </summary>
        public int MaxNonEmptyCellInEmptyRow { get; set; } = 2;

        /// <summary>
        ///     Null when created as scanning reader estimating where table ends by all cells having null value.
        /// </summary>
        public int? Count { get; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowIndex">
        ///     Relative zero-bound row index
        /// </param>
        /// <returns></returns>
        public ExcelTableRowReader this[int rowIndex]
        {
            get
            {

                if (rowIndex < 0 || rowIndex >= Count)
                    throw new IndexOutOfRangeException($"Index {rowIndex} is out of bounds (0-{Count - 1})");

                if (!_endRowIndex.HasValue && IsRowEmpty(rowIndex + _startRowIndex))
                {
                    // this is case when table boundaries are estimated based on cell contents
                    throw new IndexOutOfRangeException($"Row #{rowIndex} is empty thus considered not to belong to the table");
                }

                return new ExcelTableRowReader(ExcelWorksheet, rowIndex + _startRowIndex, _columnNameIndex);
            }
        }

        ITableRowReader ITableRowReaderCollection.this[int rowIndex] => this[rowIndex];

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerator<ITableRowReader> GetEnumerator()
        {
            var maxIndex = _endRowIndex ?? MaxExcelRowCount + 1;
            var consecutiveBlanks = 0;
            for (var rowIndex = _startRowIndex; rowIndex < maxIndex; ++rowIndex)
            {
                if (!_endRowIndex.HasValue && IsRowEmpty(rowIndex))
                {
                    if (++consecutiveBlanks > 1)
                        break;
                    continue;
                }

                consecutiveBlanks = 0;

                yield return new ExcelTableRowReader(ExcelWorksheet, rowIndex, _columnNameIndex);
            }
        }

        private bool IsRowEmpty(int oneBasedRowIndex)
        {
            var nonEmptyCellCount = _columnNameIndex.Count(p => ExcelWorksheet.Cells[oneBasedRowIndex, p.Value].Value != null);

            return nonEmptyCellCount <= MaxNonEmptyCellInEmptyRow;
        }
    }
}