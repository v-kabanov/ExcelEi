// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-03-20
// Comment		
// **********************************************************************************************/

using System;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Describes source of data for column e.g. for export into a table.
    /// </summary>
    public interface IColumnDataSource
    {
        /// <summary>
        ///     Name in the data source's data item (e.g. row or data POCO).
        /// </summary>
        string Name { get; }

        /// <summary>
        ///     Type of value contained in the column in the data source.
        /// </summary>
        Type DataType { get; }

        /// <summary>
        ///     Function accepting data item (e.g. row or data POCO) and returning column value
        /// </summary>
        Func<object, object> ValueExtractor { get; }
    }
}