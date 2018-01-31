// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-02-21
// Comment		
// **********************************************************************************************/

using System;
using AutoMapper;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Type safe value resolver for AutoMapper taking column from <see cref="ITableRowReader"/>.
    ///     Resolves value for given <typeparamref name="TE"/> property by taking it from column with given name.
    /// </summary>
    /// <typeparam name="TE">
    ///     Mapping target entity type
    /// </typeparam>
    /// <typeparam name="TV">
    ///     Mapping target value type
    /// </typeparam>
    public class RowReaderValueResolver<TE, TV> : IValueResolver<ITableRowReader, TE, TV>
    {
        public string ColumnName { get; }

        private readonly Func<ITableRowReader, TV> _valueGetter;

        public RowReaderValueResolver(string columnName)
        {
            ColumnName = columnName;
            _valueGetter = r => r.GetValue<TV>(ColumnName);
        }

        public RowReaderValueResolver(string columnName, Func<object, TV> customConverter)
        {
            Check.DoRequireArgumentNotNull(columnName, nameof(columnName));
            Check.DoRequireArgumentNotNull(customConverter, nameof(customConverter));

            ColumnName = columnName;
            _valueGetter = r => customConverter(r.GetValue<object>(ColumnName));
        }

        public TV Resolve(ITableRowReader source, TE destination, TV destMember, ResolutionContext context)
        {
            return _valueGetter(source);
        }
    }
}