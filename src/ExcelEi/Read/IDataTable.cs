// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-19
// Comment		
// **********************************************************************************************/

using System.Collections;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Enumerable collection of data objects.
    /// </summary>
    public interface IDataTable
    {
        IEnumerable Rows { get; }
    }
}