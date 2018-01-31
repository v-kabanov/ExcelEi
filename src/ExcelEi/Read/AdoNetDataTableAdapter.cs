// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-19
// Comment		
// **********************************************************************************************/

using System.Collections;
using System.Data;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Wraps ADO.NET <see cref="DataTable"/> and exposes it as <see cref="IDataTable"/>.
    /// </summary>
    public class AdoNetDataTableAdapter : IDataTable
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="AdoNetDataTableAdapter"/> class.
        /// </summary>
        public AdoNetDataTableAdapter(DataTable dataTable)
        {
            DataTable = dataTable;
        }

        public DataTable DataTable { get; }

        public IEnumerable Rows => DataTable.Rows;
    }
}