// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-03-20
// Comment		
// **********************************************************************************************/

using System.Collections.Generic;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Describes table-like data source for the purposes of extracting data from it.
    /// </summary>
    public interface ITableDataSource
    {
        /// <summary>
        ///     All columns in the data source.
        /// </summary>
        IEnumerable<IColumnDataSource> AllColumns { get; }

        /// <summary>
        ///     Get column by name
        /// </summary>
        /// <param name="name">
        /// </param>
        /// <returns>
        /// </returns>
        IColumnDataSource GetColumn(string name);
    }
}