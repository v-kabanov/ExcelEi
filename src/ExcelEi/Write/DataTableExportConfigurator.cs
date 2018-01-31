// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-23
// Comment		
// **********************************************************************************************/

using System;
using System.Data;
using ExcelEi.Read;

namespace ExcelEi.Write
{
    /// <summary>
    ///     Convenience class helping to configure export of an ADO.NET data table to excel sheet.
    /// </summary>
    public class DataTableExportConfigurator
    {
        private readonly DataTable _table;

        private DataTableExportConfigurator(DataTable table)
        {
            Check.DoRequireArgumentNotNull(table, nameof(table));

            _table = table;
            Config = new DataTableExportAutoConfig
            {
                SheetName = table.TableName,
                DataTableName = table.TableName
            };
        }

        public static DataTableExportConfigurator Begin(DataView dataView)
        {
            Check.DoRequireArgumentNotNull(dataView, "dataView");
            return new DataTableExportConfigurator(dataView.ToTable());
        }

        public static DataTableExportConfigurator Begin(DataTable dataTable)
        {
            Check.DoRequireArgumentNotNull(dataTable, "dataTable");
            return new DataTableExportConfigurator(dataTable);
        }

        /// <summary>
        ///     Export column from source data table into another sheet column.
        /// </summary>
        /// <param name="dataColumnName">
        ///     Name of the column in the data table <see cref="DataColumn.ColumnName"/>.
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Sheet column header text.
        /// </param>
        /// <returns>
        ///     Itself, fluent interface.
        /// </returns>
        public DataTableExportConfigurator AddColumn(string dataColumnName, string sheetColumnCaption)
        {
            return AddColumn(dataColumnName, Config.Columns.Count, sheetColumnCaption, null, null);
        }

        public DataTableExportConfigurator AddColumn(string dataColumnName, string sheetColumnCaption, bool autoFit)
        {
            return AddColumn(dataColumnName, Config.Columns.Count, sheetColumnCaption, autoFit, null);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataColumnName"></param>
        /// <param name="sheetColumnIndex">
        ///     0-based
        /// </param>
        /// <param name="sheetColumnCaption"></param>
        /// <param name="autoFit">
        ///     null to leave default setting (as defined by <see cref="DataColumnExportAutoConfig"/>.
        /// </param>
        /// <param name="format">
        ///     null to leave default format (as defined by <see cref="DataColumnExportAutoConfig"/>.
        ///     empty string overrides default format.
        /// </param>
        /// <returns></returns>
        public DataTableExportConfigurator AddColumn(string dataColumnName, int sheetColumnIndex, string sheetColumnCaption, bool? autoFit, string format)
        {
            Check.DoRequireArgumentNotNull(dataColumnName, "dataColumnName");
            var dataColumn = _table.Columns[dataColumnName];
            Check.DoCheckArgument(dataColumn != null, () => $"Column {dataColumnName} not found in data table");

            var columnSource = new DataColumnSource(dataColumn);
            var config = new DataColumnExportAutoConfig(Config, sheetColumnIndex, sheetColumnCaption, columnSource);

            if (autoFit.HasValue)
            {
                config.AutoFit = autoFit.Value;
            }

            if (format != null)
            {
                config.Format = format;
            }

            Config.AddColumn(config);

            return this;
        }

        /// <summary>
        ///     Safe means you do not need to worry about nulls.
        /// </summary>
        /// <param name="dataColumnName">
        ///     Name of the column in the source DataTable
        /// </param>
        /// <param name="caption">
        ///     Caption for the column in Excel.
        /// </param>
        /// <returns>
        ///     itself (fluent interface)
        /// </returns>
        public DataTableExportConfigurator AddSafeConvertingDataViewColumnInteger(string dataColumnName, string caption)
        {
            return AddSafeConvertingDataViewColumn(dataColumnName, caption, Convert.ToInt64);
        }

        /// <summary>
        ///     Safe means you do not need to worry about nulls.
        /// </summary>
        /// <param name="dataColumnName">
        ///     Name of the column in the source DataTable
        /// </param>
        /// <param name="caption">
        ///     Caption for the column in Excel.
        /// </param>
        /// <returns>
        ///     itself (fluent interface)
        /// </returns>
        public DataTableExportConfigurator AddSafeConvertingDataViewColumnDouble(string dataColumnName, string caption)
        {
            return AddSafeConvertingDataViewColumn(dataColumnName, caption, Convert.ToDouble);
        }

        /// <summary>
        ///     Add column extracting value and converting it before putting into spreadsheet.
        ///     Column is added to the end of the list and will output into next empty sheet column.
        /// </summary>
        /// <param name="dataColumnName">
        ///     Name by which source column is known in underlying data table.
        /// </param>
        /// <param name="caption">
        ///     Caption to set in Excel.
        /// </param>
        /// <param name="conversionFunction">
        ///     Function converting raw column value before outputting it to Excel. The function does not have to handle null
        ///     parameter as it will never be called for null
        /// </param>
        /// <returns></returns>
        public DataTableExportConfigurator AddSafeConvertingDataViewColumn<T>(string dataColumnName, string caption, Func<object, T> conversionFunction)
        {
            Check.DoRequireArgumentNotNull(dataColumnName, "dataColumnName");
            Check.DoRequireArgumentNotNull(caption, "caption");
            Check.DoRequireArgumentNotNull(conversionFunction, "conversionFunction");

            var columnSource = new DataColumnSource(dataColumnName, typeof(T), r => GetDataRowColumnValue(r, dataColumnName, conversionFunction));

            var columnConfig = new DataColumnExportAutoConfig(Config, Config.Columns.Count, caption, columnSource);

            Config.AddColumn(columnConfig);

            return this;
        }

        /// <summary>
        ///     Configuration being constructed.
        /// </summary>
        public DataTableExportAutoConfig Config { get; }

        private static object GetDataRowColumnValue<T>(object rowObject, string columnName, Func<object, T> conversionFunction)
        {
            object result = null;

            var rawValue = ((DataRow)rowObject)[columnName];

            if (rawValue != null && !Convert.IsDBNull(rawValue))
            {
                result = conversionFunction(rawValue);
            }

            return result;
        }
    }
}