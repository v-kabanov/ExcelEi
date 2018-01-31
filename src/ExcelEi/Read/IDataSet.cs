// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-19
// Comment		
// **********************************************************************************************/

using System.Collections.Generic;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Collection of named <see cref="IDataTable"/>.
    /// </summary>
    public interface IDataSet
    {
        IDictionary<string, IDataTable> DataTables { get; }
    }
}