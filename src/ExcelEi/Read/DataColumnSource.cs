// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-03-20
// Comment		
// **********************************************************************************************/

using System;
using System.Data;
using System.Data.SqlTypes;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Describes <see cref="DataColumn"/> in ADO.NET's <see cref="DataTable"/> for the purposes of extracting data from its <see cref="DataRow"/>s.
    /// </summary>
    public class DataColumnSource : IColumnDataSource
    {
        public DataColumnSource(DataColumn dataColumn)
        {
            Check.DoRequireArgumentNotNull(dataColumn, nameof(dataColumn));
            Check.DoCheckArgument(dataColumn.Table != null, "Column must belong to a table");

            Name = dataColumn.ColumnName;
            DataType = dataColumn.DataType;

            if (DataType == typeof(SqlDateTime))
            {
                ValueExtractor = o => GetDateTimeValue(((DataRow)o)[Name]);
            }
            else
            {
                ValueExtractor = o => ((DataRow)o)[Name];
            }
        }

        public DataColumnSource(string name, Type dataType, Func<object, object> valueExtractor)
        {
            Check.DoRequireArgumentNotNull(name, nameof(name));
            Check.DoRequireArgumentNotNull(dataType, nameof(dataType));
            Check.DoRequireArgumentNotNull(valueExtractor, nameof(valueExtractor));

            Name = name;
            DataType = dataType;
            ValueExtractor = valueExtractor;
        }

        public string Name { get; }

        public Type DataType { get; }

        /// <inheritdoc />
        public object GetValue(object dataObject)
        {
            return ValueExtractor(dataObject);
        }

        public Func<object, object> ValueExtractor { get; }

        public DateTime? GetDateTimeValue(object value)
        {
            if (value == null) return null;

            if (value is SqlDateTime)
                return ((SqlDateTime)value).Value;

            return value as DateTime?;
        }
    }
}