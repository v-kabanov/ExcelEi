// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-19
// Comment		
// **********************************************************************************************/

using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;

namespace ExcelEi.Write
{
    /// <summary>
    ///     Creates default configuration from existing ADO.NET <see cref="DataSet"/>.
    /// </summary>
    public class DataSetExportAutoConfig : IWorkbookExportConfig
    {
        private readonly List<ISheetExportConfig> _tables = new List<ISheetExportConfig>();

        public DataSetExportAutoConfig()
        {
            SheetTables = new ReadOnlyCollection<ISheetExportConfig>(_tables);
        }

#if !NOADONET

        /// <summary>
        ///     Initializes a new instance of the <see cref="DataSetExportAutoConfig"/> class.
        /// </summary>
        public DataSetExportAutoConfig(DataSet dataSet)
            : this()
        {
            Check.DoRequireArgumentNotNull(dataSet, "dataSet");

            _tables.Capacity = dataSet.Tables.Count;

            foreach (DataTable table in dataSet.Tables)
            {
                _tables.Add(new DataTableExportAutoConfig(table));
            }
        }

        /// <summary>
        ///     Initializes new instance with single ADO.NET <see cref="DataView"/> instance from which to create
        ///     default export configuration.
        /// </summary>
        /// <param name="dataView">
        ///     Data view produced by e.g. ASP.NET SqlDataSource.
        /// </param>
        public DataSetExportAutoConfig(DataView dataView)
            : this(dataView.Table)
        {
        }

        /// <summary>
        ///     Initializes new instance with single ADO.NET <see cref="DataTable"/> instance from which to create
        ///     default export configuration.
        /// </summary>
        /// <param name="dataTable">
        ///     Data table.
        /// </param>
        public DataSetExportAutoConfig(DataTable dataTable)
            : this()
        {
            Check.DoRequireArgumentNotNull(dataTable, "dataTable");
            _tables.Add(new DataTableExportAutoConfig(dataTable));
        }
#endif

        /// <summary>
        ///     Initializes new instance with single provided table/sheet export config.
        /// </summary>
        /// <param name="sheetConfig">
        ///     Defines export of 1 data table into 1 sheet.
        /// </param>
        public DataSetExportAutoConfig(ISheetExportConfig sheetConfig)
            : this()
        {
            Check.DoRequireArgumentNotNull(sheetConfig, "sheetConfig");

            _tables.Add(sheetConfig);
        }

        public void AddSheet(ISheetExportConfig sheetConfig)
        {
            Check.DoRequireArgumentNotNull(sheetConfig, "sheetConfig");

            _tables.Add(sheetConfig);
        }

        /// <summary>
        ///     Get data table config for further customisation.
        /// </summary>
        /// <param name="sheetName">
        ///     Case sensitive, to be matched with <see cref="ISheetExportConfig.SheetName"/>
        /// </param>
        /// <returns></returns>
        public DataTableExportAutoConfig GetTableConfig(string sheetName)
        {
            return (DataTableExportAutoConfig)SheetTables.FirstOrDefault(t => t.SheetName == sheetName);
        }

        /// <inheritdoc />
        public IList<ISheetExportConfig> SheetTables { get; }
    }
}