// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-03-20
// Comment		
// **********************************************************************************************/

using System.Collections.Generic;
using System.Data;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Describes ADO.NET <see cref="DataTable"/> for the purposes of extracting data from it.
    /// </summary>
    public class DataTableSource : ITableDataSource
    {
        public DataTable DataTable { get; }

        public DataTableSource(DataTable dataTable)
        {
            Check.DoRequireArgumentNotNull(dataTable, nameof(dataTable));

            DataTable = dataTable;
        }

        public IEnumerable<IColumnDataSource> AllColumns
        {
            get
            {
                for (var i = 0; i < DataTable.Columns.Count; ++i)
                    yield return new DataColumnSource(DataTable.Columns[i]);
            }
        }

        public IColumnDataSource GetColumn(string name)
        {
            Check.DoRequireArgumentNotNull(name, nameof(name));
            Check.DoCheckArgument(DataTable.Columns.Contains(name), () => $"Column {name} not found");

            return new DataColumnSource(DataTable.Columns[name]);
        }
    }
}