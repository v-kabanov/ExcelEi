// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-19
// Comment		
// **********************************************************************************************/

using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Proxy exposing various data sources as collection of data tables with unique names - abstraction between e.g. exporter and data.
    /// </summary>
    public class DataSetAdapter : IDataSet
    {
        private readonly Dictionary<string, IDataTable> _tables;

        /// <summary>
        ///     Initializes new instance with ADO.NET data set.
        /// </summary>
        public DataSetAdapter(DataSet dataSet)
            : this()
        {
            Check.DoRequireArgumentNotNull(dataSet, nameof(dataSet));

            foreach (DataTable table in dataSet.Tables)
            {
                _tables.Add(table.TableName, new AdoNetDataTableAdapter(table));
            }
        }

        /// <summary>
        ///     Initializes new instance with single ADO.NET data view, received e.g. from ASP.NET SqlDataSource./>.
        /// </summary>
        public DataSetAdapter(DataView dataView)
            : this()
        {
            var dataTable = dataView.ToTable();
            _tables.Add(dataTable.TableName, new AdoNetDataTableAdapter(dataTable));
        }

        public DataSetAdapter()
        {
            _tables = new Dictionary<string, IDataTable>();
        }

        /// <summary>
        ///     Add new table
        /// </summary>
        /// <param name="dataTable">
        ///     Mandatory
        /// </param>
        /// <param name="name">
        ///     Mandatory, must be unique
        /// </param>
        /// <returns>
        ///     Itself for chaining
        /// </returns>
        public DataSetAdapter Add(IDataTable dataTable, string name)
        {
            Check.DoRequireArgumentNotNull(dataTable, nameof(dataTable));
            Check.DoRequireArgumentNotNull(name, nameof(name));
            Check.DoCheckArgument(!_tables.ContainsKey(name), () => $"Table {name} already exists");

            _tables.Add(name, dataTable);
            return this;
        }

        /// <summary>
        ///     Add new table from POCO collection
        /// </summary>
        /// <typeparam name="T">
        ///     POCO type
        /// </typeparam>
        /// <param name="pocoCollection">
        ///     Mandatory
        /// </param>
        /// <param name="name">
        ///     Mandatory, must be unique
        /// </param>
        /// <returns>
        ///     Itself for chaining
        /// </returns>
        public DataSetAdapter Add<T>(IEnumerable<T> pocoCollection, string name)
        {
            return Add(new PocoTableAdapter(pocoCollection), name);
        }

        /// <summary>
        ///     Add new table from POCO collection
        /// </summary>
        /// <param name="pocoCollection">
        ///     Mandatory
        /// </param>
        /// <param name="name">
        ///     Mandatory, must be unique
        /// </param>
        /// <returns>
        ///     Itself for chaining
        /// </returns>
        public DataSetAdapter Add(IEnumerable pocoCollection, string name)
        {
            return Add(new PocoTableAdapter(pocoCollection), name);
        }

        public IDictionary<string, IDataTable> DataTables => _tables;
    }
}