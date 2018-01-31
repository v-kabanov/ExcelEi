// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-01-17
// Comment		
// **********************************************************************************************/

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Reader for EPPlus excel table. EPPlus has the advandage of preserving cell data types.
    /// </summary>
    public class ExcelTableReader : ITableReader
    {
        // column names in original order
        private readonly IList<string> _columnNames;

        /// <param name="excelTable">
        ///     Mandatory
        /// </param>
        public ExcelTableReader(ExcelTable excelTable)
        {
            Check.DoRequireArgumentNotNull(excelTable, nameof(excelTable));

            _columnNames = excelTable.Columns.Select(c => c.Name).ToList();
            var columnNameIndex = excelTable.Columns.ToDictionary(c => c.Name, c => c.Position + excelTable.Address.Start.Column);

            var startRow = excelTable.Address.Start.Row;
            if (excelTable.ShowHeader)
                ++startRow;

            Rows = new ExcelTableRowReaderCollection(startRow, excelTable.Address.End.Row, excelTable.WorkSheet, columnNameIndex);
        }

        /// <summary>
        ///     Read virtual table from epplus worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="startRowIndex">
        ///     1-based index (1..1048576), inclusive
        /// </param>
        /// <param name="endRowIndex">
        ///     Optional 1-based index (1..1048577), exclusive to allow reading empty table
        /// </param>
        /// <param name="columns">
        ///     List of columns in original order
        /// </param>
        public ExcelTableReader(ExcelWorksheet worksheet, int startRowIndex, int? endRowIndex, IList<KeyValuePair<string, int>> columns)
        {
            Check.DoRequireArgumentNotNull(worksheet, nameof(worksheet));
            Check.DoRequireArgumentNotNull(columns, nameof(columns));

            _columnNames = columns.Select(p => p.Key).ToList();
            var columnNameIndex = columns.ToDictionary(p => p.Key, p => p.Value);

            Rows = new ExcelTableRowReaderCollection(startRowIndex, endRowIndex, worksheet, columnNameIndex);
        }

        public ExcelTableRowReaderCollection Rows { get; }

        ITableRowReaderCollection ITableReader.Rows => Rows;

        /// <summary>
        ///     Columns in original order as they appear in worksheet.
        /// </summary>
        public ICollection<string> Columns => _columnNames;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="headerRowIndex">
        ///     1-based index (1..1048576)
        /// </param>
        /// <returns></returns>
        public static ExcelTableReader ReadContiguousTableWithHeader(ExcelWorksheet worksheet, int headerRowIndex)
        {
            Check.DoRequireArgumentNotNull(worksheet, nameof(worksheet));

            var fullRow = worksheet.Cells[headerRowIndex, 1, headerRowIndex, ExcelPackage.MaxColumns];
            var firstColumn = fullRow.FirstOrDefault(r => r.Value != null);
            Check.DoCheckArgument(firstColumn != null, "Table header not found");
            Debug.Assert(firstColumn != null);

            // last column could be same as start column
            var headerCells = fullRow
                .Where(r => r.Start.Column >= firstColumn.Start.Column)
                .TakeWhile(r => r.Value != null)
                .ToList();
            Check.DoCheckArgument(headerCells.Count > 0, "No columns in header");

            Check.DoCheckArgument(headerCells.All(c => c.Value is string), "Header cells contain non-text values");

            var columns = headerCells.Select(c => new KeyValuePair<string, int>((string)c.Value, c.Start.Column)).ToList();
            // two 1-based indexes
            var startRowIndexInclusive = headerRowIndex + 1;

            return new ExcelTableReader(worksheet, startRowIndexInclusive, null, columns);
        }

        /// <summary>
        ///     Read arbitrary table, possibly without column headers
        /// </summary>
        /// <param name="worksheet">
        ///     Mandatory
        /// </param>
        /// <param name="startRowIndex">
        ///     1-based index of first row with data, inclusive
        /// </param>
        /// <param name="endRowIndexExclusive">
        ///     1-based index of last row with data, exclusive to allow reading empty tables
        /// </param>
        /// <param name="columns">
        ///     Defines columns, possibly sparse
        /// </param>
        /// <returns>
        ///     Reader will return even rows with all cells empty, reading whole range defined by parameters
        /// </returns>
        public static ExcelTableReader ReadArbitraryTable(ExcelWorksheet worksheet, int startRowIndex, int endRowIndexExclusive, IList<KeyValuePair<string, int>> columns)
        {
            Check.DoRequireArgumentNotNull(worksheet, nameof(worksheet));
            Check.DoRequireArgumentNotNull(columns, nameof(columns));
            Check.DoCheckArgument(columns.Count > 0, "No columns in table to read");

            return new ExcelTableReader(worksheet, startRowIndex, endRowIndexExclusive, columns);
        }
    }
}