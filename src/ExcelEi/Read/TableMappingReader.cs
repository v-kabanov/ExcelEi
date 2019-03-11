// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-02-21
// Comment		
// **********************************************************************************************/

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading;
using AutoMapper;
using AutoMapper.Configuration;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Simplifies the use of automapper for mapping excel table rows domain objects.
    ///     Keeps mapping locally bypassing global AutoMapper configurations, allowing to map same entities differently in the same application
    ///     which is not primary mode for AutoMapper.
    /// </summary>
    /// <typeparam name="T">
    ///     Type of entities being read.
    /// </typeparam>
    /// <remarks>
    ///     Mapping is relatively expensive, so instances of this class should be kept in a static field.
    /// </remarks>
    public class TableMappingReader<T>
    {
        private readonly Lazy<IMapper> _mapper; // MapperConfiguration _mapperConfiguration;
        private readonly IMappingExpression<ITableRowReader, T> _mappingExpression;

        private readonly IList<KeyValuePair<MemberInfo, string>> _mappedMembers;

        public TableMappingReader()
        {
            var configurationExpression = new MapperConfigurationExpression();
            _mappingExpression = configurationExpression.CreateMap<ITableRowReader, T>();

            _mapper = new Lazy<IMapper>(() => new Mapper(new MapperConfiguration(configurationExpression)), LazyThreadSafetyMode.ExecutionAndPublication);

            _mappedMembers = new List<KeyValuePair<MemberInfo, string>>();

            MappedMembers = new ReadOnlyCollection<KeyValuePair<MemberInfo, string>>(_mappedMembers);
        }

        /// <summary>
        ///     List contains key-value pairs with details of members (properties or fields) mapped so far. Non-member mappings are not included.
        /// </summary>
        public IReadOnlyCollection<KeyValuePair<MemberInfo, string>> MappedMembers { get; }

        public IMapper Mapper => _mapper.Value;

        public TableMappingReader<T> Map<V>(Expression<Func<T, V>> memberReference)
        {
            var memberInfo = ExpressionHelper.GetMember(memberReference);

            Check.DoRequire(memberInfo != null, "Cannot resolve member reference");
            Debug.Assert(memberInfo != null, "memberInfo != null");

            return Map(memberReference, memberInfo.Name);
        }

        public TableMappingReader<T> Map<V>(Expression<Func<T, V>> propertyReference, string columnName)
        {
            Check.DoRequireArgumentNotNull(propertyReference, nameof(propertyReference));

            _mappingExpression.ForMember(propertyReference, opt => opt.MapFrom(new RowReaderValueResolver<T, V>(columnName)));

            RegisterMemberMapping(propertyReference, columnName);

            return this;
        }

        public TableMappingReader<T> Map<V>(Expression<Func<T, V>> propertyReference, string columnName, Func<object, V> customConverter)
        {
            Check.DoRequireArgumentNotNull(propertyReference, nameof(propertyReference));

            _mappingExpression.ForMember(propertyReference, opt => opt.MapFrom(new RowReaderValueResolver<T, V>(columnName, customConverter)));

            RegisterMemberMapping(propertyReference, columnName);

            return this;
        }

        /// <summary>
        ///     Read list of entity instances from raw table reader.
        /// </summary>
        /// <param name="tableReader"></param>
        /// <returns></returns>
        public IList<T> Read(ITableReader tableReader)
        {
            Check.DoRequireArgumentNotNull(tableReader, nameof(tableReader));

            return Mapper.Map<IEnumerable<ITableRowReader>, IList<T>>(tableReader.Rows);
        }

        /// <summary>
        ///     Read entity from single row.
        /// </summary>
        /// <param name="rowReader">
        ///     Mandatory
        /// </param>
        /// <returns>
        ///     New entity instance
        /// </returns>
        public T Read(ITableRowReader rowReader)
        {
            Check.DoRequireArgumentNotNull(rowReader, nameof(rowReader));

            return Mapper.Map<ITableRowReader, T>(rowReader);
        }

        /// <summary>
        ///     Read virtual table (region) from excel worksheet, identifying columns by their headings in the known row.
        /// </summary>
        /// <param name="path">
        ///     Mandatory
        /// </param>
        /// <param name="sheetName">
        ///     Optional, defaults to first one
        /// </param>
        /// <param name="headerRowIndex">
        ///     1-based index
        /// </param>
        public IList<T> ReadContiguousExcelTableWithHeader(string path, string sheetName, int headerRowIndex)
        {
            Check.DoRequireArgumentNotNull(path, nameof(path));

            var reader = AdoTableReader.ReadContiguousExcelTableWithHeader(path, sheetName, headerRowIndex);

            return Read(reader);
        }

        /// <summary>
        ///     Read arbitrary table, possibly without column headers
        /// </summary>
        /// <param name="path">
        ///     Mandatory
        /// </param>
        /// <param name="sheetName">
        ///     Optional, defaults to first one
        /// </param>
        /// <param name="startRowIndex">
        ///     1-based index of first row with data, inclusive
        /// </param>
        /// <param name="endRowIndexExclusive">
        ///     1-based index of last row with data, exclusive to allow reading empty tables
        /// </param>
        /// <param name="columns">
        ///     Defines columns with 1-based indexes, possibly sparse
        /// </param>
        public IList<T> ReadArbitraryExcelTable(
            string path, string sheetName, int startRowIndex, int endRowIndexExclusive, IList<KeyValuePair<string, int>> columns)
        {
            var reader = AdoTableReader.ReadArbitraryExcelTable(path, sheetName, startRowIndex, endRowIndexExclusive, columns);

            return Read(reader);
        }

        private void RegisterMemberMapping(LambdaExpression expression, string columnName)
        {
            var memberInfo = ExpressionHelper.GetMember(expression);
            if (null != memberInfo)
                _mappedMembers.Add(new KeyValuePair<MemberInfo, string>(memberInfo, columnName));
        }
    }
}