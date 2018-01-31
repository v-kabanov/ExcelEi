// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-04-04
// Comment		
// **********************************************************************************************/

using System;
using System.Collections.Generic;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Abstract row with named columns available for reading
    /// </summary>
    public interface ITableRowReader
    {
        IEnumerable<string> AllColumnNames { get; }

        /// <summary>
        ///     Get raw column value.
        /// </summary>
        /// <param name="columnName">
        ///     Mandatory, column must exist, otherwise exception is thrown
        /// </param>
        /// <exception cref="ArgumentException">
        ///     Column with given <paramref name="columnName"/> is not found.
        /// </exception>
        object this[string columnName] { get; }

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
        T GetValue<T>(string columnName);
    }
}