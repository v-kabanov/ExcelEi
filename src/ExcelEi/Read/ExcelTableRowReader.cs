// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-01-17
// Comment		
// **********************************************************************************************/

using System;
using System.Collections.Generic;
using OfficeOpenXml;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Row reader for EPPlus.
    /// </summary>
    public class ExcelTableRowReader : ITableRowReader
    {
        public ExcelWorksheet ExcelWorksheet { get; }

        /// <summary>
        ///     1-based index in <see cref="ExcelWorksheet"/> of the row to read
        /// </summary>
        public int RowIndex { get; }

        /// <summary>
        ///     Index resolves name to <see cref="ExcelWorksheet"/> column index which can be passed to <see cref="ExcelWorksheet.Cells"/>
        /// </summary>
        private readonly IDictionary<string, int> _columnNameIndex;

        /// <param name="excelWorksheet"></param>
        /// <param name="rowIndex">
        ///     1-based index
        /// </param>
        /// <param name="columnNameIndex">
        ///     Maps name to 1-based worksheet column index
        /// </param>
        public ExcelTableRowReader(ExcelWorksheet excelWorksheet, int rowIndex, IDictionary<string, int> columnNameIndex)
        {
            Check.DoRequireArgumentNotNull(excelWorksheet, nameof(excelWorksheet));
            Check.DoRequireArgumentNotNull(columnNameIndex, nameof(columnNameIndex));
            Check.DoCheckArgument(rowIndex > 0 && rowIndex <= ExcelTableRowReaderCollection.MaxExcelRowCount
                , () => $"Row index out of bounds: {rowIndex}");

            ExcelWorksheet = excelWorksheet;
            RowIndex = rowIndex;
            _columnNameIndex = columnNameIndex;
        }

        public IEnumerable<string> AllColumnNames => _columnNameIndex.Keys;

        /// <summary>
        ///     Get raw column value.
        /// </summary>
        /// <param name="columnName">
        ///     Mandatory, column must exist, otherwise exception is thrown
        /// </param>
        /// <exception cref="ArgumentException">
        ///     Column with given <paramref name="columnName"/> is not found.
        /// </exception>
        public object this[string columnName] => ExcelWorksheet.Cells[RowIndex, GetColumnIndex(columnName)].Value;

        /// <summary>
        ///     Get strongly typed column value. See <see cref="Conversion.GetTypedExcelValue{T}"/> for conversion rules.
        /// </summary>
        /// <typeparam name="T">
        ///     Data type to return
        /// </typeparam>
        /// <param name="columnName">
        ///     Mandatory, column must exist, otherwise exception is thrown
        /// </param>
        /// <exception cref="ArgumentException">
        ///     Column with given <paramref name="columnName"/> is not found.
        /// </exception>
        public T GetValue<T>(string columnName)
        {
            // the GetTypesValue method contains bugs; cannot convert string to integer;
            // return ExcelWorksheet.Cells[RowIndex, GetColumnIndex(columnName)].GetValue<T>();
            var val = this[columnName];
            return Conversion.GetTypedExcelValue<T>(val);
        }

        private int GetColumnIndex(string columnName)
        {
            int result;
            if (_columnNameIndex.TryGetValue(columnName, out result))
                return result;

            throw new ArgumentException($"Unknown column: {columnName}");
        }
    }
}