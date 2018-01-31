// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-04-04
// Comment		
// **********************************************************************************************/

using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelEi.Read
{
    public class AdoDataTableRowReader : ITableRowReader
    {
        public DataTable DataTable => DataRow.Table;

        private readonly IDictionary<string, int> _columnNameIndex;

        /// <param name="dataRow">
        /// </param>
        /// <param name="columnNameIndex">
        ///     Maps name to 0-based table column index
        /// </param>
        public AdoDataTableRowReader(DataRow dataRow, IDictionary<string, int> columnNameIndex)
        {
            Check.DoRequireArgumentNotNull(dataRow, nameof(dataRow));
            Check.DoRequireArgumentNotNull(columnNameIndex, nameof(columnNameIndex));

            DataRow = dataRow;

            _columnNameIndex = columnNameIndex;
        }

        public DataRow DataRow { get; }

        public IEnumerable<string> AllColumnNames => _columnNameIndex.Keys;

        /// <inheritdoc />
        public object this[string columnName]
        {
            get
            {
                var columnIndex = GetColumnIndex(columnName);
                var rawValue = DataRow[columnIndex];
                return IsNull(rawValue) ? null : rawValue;
            }
        }

        /// <inheritdoc />
        public T GetValue<T>(string columnName)
        {
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

        /// <summary>
        ///     Check if value is null from the point of view of ADO.NET.
        /// </summary>
        public static bool IsNull(object rawValue)
        {
            return rawValue == null || rawValue == DBNull.Value;
        }
    }
}