// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-18
// Comment		
// **********************************************************************************************/

using System.Collections.Generic;

namespace ExcelEi.Write
{
    /// <summary>
    ///     Interface of an object defining how to export a collection of tables into [Excel] spreadsheet
    /// </summary>
    public interface IWorkbookExportConfig
    {
        /// <summary>
        ///     Read-only collection of tables to export to respective sheets keyed by sheet name.
        /// </summary>
        IList<ISheetExportConfig> SheetTables { get; }
    }
}