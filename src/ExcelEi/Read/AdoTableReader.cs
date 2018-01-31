// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-04-04
// Comment		
// **********************************************************************************************/

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ExcelDataReader;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Reads <see cref="DataTable"/> or virtual tables inside it.
    /// </summary>
    public class AdoTableReader : ITableReader
    {
        // column names in original order
        private readonly IList<string> _columnNames;

        /// <summary>
        ///     Read whole table, without renaming or excluding columns.
        /// </summary>
        /// <param name="dataTable">
        ///     Mandatory
        /// </param>
        public AdoTableReader(DataTable dataTable)
        {
            Check.DoRequireArgumentNotNull(dataTable, nameof(dataTable));

            _columnNames = dataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();

            var columnNameIndex = dataTable.Columns.Cast<DataColumn>().ToDictionary(c => c.ColumnName, c => c.Ordinal);

            Rows = new AdoTableRowReaderCollection(0, null, dataTable, columnNameIndex);
        }

        /// <summary>
        ///     Read virtual table inside ado.net <paramref name="dataTable"/>, assigning given column names;
        ///     columns of the virtual table can be sparse.
        /// </summary>
        /// <param name="dataTable">
        ///     Mandatory
        /// </param>
        /// <param name="startRowIndex">
        ///     0-based row index of the first row with data
        /// </param>
        /// <param name="endRowIndex">
        ///     Optional
        /// </param>
        /// <param name="columns">
        ///     Dictionary with key being assigned column name and value - zero-based column index in <paramref name="dataTable"/>.
        /// </param>
        public AdoTableReader(DataTable dataTable, int startRowIndex, int? endRowIndex, IList<KeyValuePair<string, int>> columns)
        {
            Check.DoRequireArgumentNotNull(dataTable, nameof(dataTable));
            Check.DoRequireArgumentNotNull(columns, nameof(columns));

            _columnNames = columns.Select(p => p.Key).ToList();
            var columnNameIndex = columns.ToDictionary(p => p.Key, p => p.Value);

            Rows = new AdoTableRowReaderCollection(startRowIndex, endRowIndex, dataTable, columnNameIndex);
        }

        public ITableRowReaderCollection Rows { get; }

        public ICollection<string> Columns => _columnNames;

        /// <summary>
        ///     
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="headerRowIndex">
        ///     0-based index
        /// </param>
        /// <returns></returns>
        public static AdoTableReader ReadContiguousTableWithHeader(DataTable dataTable, int headerRowIndex)
        {
            Check.DoRequireArgumentNotNull(dataTable, nameof(dataTable));

            var headerRow = dataTable.Rows[headerRowIndex];
            var firstColumn = dataTable.Columns.Cast<DataColumn>().FirstOrDefault(c => !AdoDataTableRowReader.IsNull(headerRow[c]));
            Check.DoCheckArgument(firstColumn != null, "Table header not found");

            Debug.Assert(firstColumn != null, nameof(firstColumn) + " != null");

            // last column could be same as start column
            var headerColumns = dataTable.Columns
                .Cast<DataColumn>()
                .Where(c => c.Ordinal >= firstColumn.Ordinal)
                .TakeWhile(c => !AdoDataTableRowReader.IsNull(headerRow[c]))
                .ToList();
            Check.DoCheckArgument(headerColumns.Count > 0, "No columns in header");

            Check.DoCheckArgument(headerColumns.All(c => headerRow[c] is string), "Header cells contain non-text values");

            var columns = headerColumns.Select(c => new KeyValuePair<string, int>((string)headerRow[c], c.Ordinal)).ToList();
            // two 1-based indexes
            var startRowIndexInclusive = headerRowIndex + 1;

            return new AdoTableReader(dataTable, startRowIndexInclusive, null, columns);
        }

        /// <summary>
        ///     Read arbitrary table, possibly without column headers
        /// </summary>
        /// <param name="dataTable">
        ///     Mandatory
        /// </param>
        /// <param name="startRowIndex">
        ///     0-based index of first row with data, inclusive
        /// </param>
        /// <param name="endRowIndexExclusive">
        ///     0-based index of last row with data, exclusive to allow reading empty tables
        /// </param>
        /// <param name="columns">
        ///     Defines columns with 0-based indexes, possibly sparse
        /// </param>
        /// <returns>
        ///     Reader will return even rows with all cells empty, reading whole range defined by parameters
        /// </returns>
        public static AdoTableReader ReadArbitraryTable(DataTable dataTable, int startRowIndex, int endRowIndexExclusive, IList<KeyValuePair<string, int>> columns)
        {
            Check.DoRequireArgumentNotNull(dataTable, nameof(dataTable));
            Check.DoRequireArgumentNotNull(columns, nameof(columns));
            Check.DoCheckArgument(columns.Count > 0, "No columns in table to read");

            return new AdoTableReader(dataTable, startRowIndex, endRowIndexExclusive, columns);
        }

        /// <summary>
        ///     Read virtual table (region) from excel worksheet, identifying columns by their headings in the known row.
        /// </summary>
        /// <param name="path">
        ///     Mandatory
        /// </param>
        /// <param name="sheetName">
        ///     Optional, defaults to first one
        /// </param>
        /// <param name="headerRowIndex">
        ///     1-based index
        /// </param>
        /// <returns></returns>
        public static AdoTableReader ReadContiguousExcelTableWithHeader(string path, string sheetName, int headerRowIndex)
        {
            Check.DoRequireArgumentNotNull(path, nameof(path));

            var table = GetWorksheetDataTable(path, sheetName);

            return ReadContiguousTableWithHeader(table, headerRowIndex - 1);
        }

        /// <summary>
        ///     Read arbitrary table, possibly without column headers
        /// </summary>
        /// <param name="path">
        ///     Mandatory
        /// </param>
        /// <param name="sheetName">
        ///     Optional, defaults to first one
        /// </param>
        /// <param name="startRowIndex">
        ///     1-based index of first row with data, inclusive
        /// </param>
        /// <param name="endRowIndexExclusive">
        ///     1-based index of last row with data, exclusive to allow reading empty tables
        /// </param>
        /// <param name="columns">
        ///     Defines columns with 1-based indexes, possibly sparse
        /// </param>
        /// <returns>
        ///     Reader will return even rows with all cells empty, reading whole range defined by parameters
        /// </returns>
        public static AdoTableReader ReadArbitraryExcelTable(
            string path, string sheetName, int startRowIndex, int endRowIndexExclusive, IList<KeyValuePair<string, int>> columns)
        {
            Check.DoRequireArgumentNotNull(path, nameof(path));

            var table = GetWorksheetDataTable(path, sheetName);

            var zeroBaseIndexColumns = columns.Select(p => new KeyValuePair<string, int>(p.Key, p.Value - 1)).ToList();

            return ReadArbitraryTable(table, startRowIndex - 1, endRowIndexExclusive - 1, zeroBaseIndexColumns);
        }

        /// <summary>
        ///     Get worksheet as data table. Supports xls and xlsx.
        /// </summary>
        /// <param name="path">
        ///     Mandatory, excel file path
        /// </param>
        /// <param name="sheetName">
        ///     Optional, defaults to first one
        /// </param>
        /// <returns></returns>
        public static DataTable GetWorksheetDataTable(string path, string sheetName = null)
        {
            Check.DoRequireArgumentNotNull(path, nameof(path));

            var dataSet = GetWorkbookDataSet(path);

            return string.IsNullOrEmpty(sheetName)
                ? dataSet.Tables[0]
                : dataSet.Tables[sheetName];
        }

        /// <summary>
        ///     Read workbook and create dataset with table per sheet. Supports xls and xlsx.
        /// </summary>
        /// <param name="path">
        ///     Mandatory
        /// </param>
        /// <returns>
        ///     not null; no need to dispose, already disposed, see remarks
        /// </returns>
        /// <remarks>
        ///     Note that DataSet's Dispose implementation is controversial and inconsistent, it and contained tables inherit IDisposable
        ///     but DataSet itself does not dispose contained tables.
        ///     See e.g. discussion at http://stackoverflow.com/questions/913228/should-i-dispose-dataset-and-datatable .
        ///     End result: we let ExcelDataReader dispose DataSet but continue to use it.
        /// </remarks>
        public static DataSet GetWorkbookDataSet(string path)
        {
            Check.DoRequireArgumentNotNull(path, nameof(path));

            var isXls = ".xls".Equals(Path.GetExtension(path), StringComparison.OrdinalIgnoreCase);
            using (var stream = File.OpenRead(path))
            using (var reader = isXls
                        ? ExcelReaderFactory.CreateBinaryReader(stream)
                        : ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                return reader.AsDataSet();
            }

        }
    }
}