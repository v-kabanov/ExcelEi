// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-04-04
// Comment		
// **********************************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelEi.Read
{
    public class AdoTableRowReaderCollection : ITableRowReaderCollection
    {
        private readonly int _startRowIndex;
        private readonly int? _endRowIndex;
        /// <summary>
        ///     Index resolves name to <see cref="DataTable"/> column index which can be passed to <see cref="DataRow"/>
        /// </summary>
        private readonly IDictionary<string, int> _columnNameIndex;

        public DataTable DataTable { get; }

        /// <param name="startRowIndex">
        ///     0-based index as understood by <see cref="this[int]"/>, inclusive.
        /// </param>
        /// <param name="endRowIndex">
        ///     Optional, 1-based index as understood by <see cref="DataRowCollection"/>, exclusive to be able to reference empty tables.
        /// </param>
        /// <param name="dataTable">
        ///     Table with data (located in sub-range of cells)
        /// </param>
        /// <param name="columnNameIndex">
        ///     Mandatory, maps column names to indexec as understood by <see cref="DataColumnCollection"/>.
        /// </param>
        public AdoTableRowReaderCollection(int startRowIndex, int? endRowIndex, DataTable dataTable, IDictionary<string, int> columnNameIndex)
        {
            Check.DoRequireArgumentNotNull(dataTable, nameof(dataTable));
            Check.DoRequireArgumentNotNull(columnNameIndex, nameof(columnNameIndex));
            Check.DoCheckArgument(startRowIndex >= 0, "Invalid row index, must be positive");
            Check.DoCheckArgument(!endRowIndex.HasValue || (endRowIndex >= startRowIndex && endRowIndex < dataTable.Rows.Count)
                , "Invalid row range");
            Check.DoCheckArgument(!columnNameIndex.Any(p => p.Value < 0 || p.Value >= dataTable.Columns.Count)
                , "Column index invalid (out of bounds)");


            _startRowIndex = startRowIndex;
            _endRowIndex = endRowIndex;
            DataTable = dataTable;
            _columnNameIndex = columnNameIndex;

            if (_endRowIndex.HasValue)
                Count = _endRowIndex - _startRowIndex;
        }

        /// <summary>
        ///     Get or set max number of cells in the row which still allow to consider it empty.
        /// </summary>
        public int MaxNonEmptyCellInEmptyRow { get; set; } = 2;

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerator<ITableRowReader> GetEnumerator()
        {
            var maxIndex = _endRowIndex ?? DataTable.Rows.Count;
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

                yield return new AdoDataTableRowReader(DataTable.Rows[rowIndex], _columnNameIndex);
            }
        }

        public int? Count { get; }
        
        public ITableRowReader this[int rowIndex]
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

                return new AdoDataTableRowReader(DataTable.Rows[rowIndex], _columnNameIndex);
            }
        }

        private bool IsRowEmpty(int zeroBasedRowIndex)
        {
            var nonEmptyCellCount = _columnNameIndex.Count(p => !AdoDataTableRowReader.IsNull(DataTable.Rows[zeroBasedRowIndex][p.Value]));

            return nonEmptyCellCount <= MaxNonEmptyCellInEmptyRow;
        }
    }
}