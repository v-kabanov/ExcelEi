// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-03-20
// Comment		
// **********************************************************************************************/

using System;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Defines column values source taking them from POCOs.
    /// </summary>
    /// <typeparam name="TA">
    ///     POCO type; if an instance presented for extraction is not <typeparamref name="TA"/>, null is passed to
    ///     <see cref="ValueExtractor"/>.
    /// </typeparam>
    /// <typeparam name="TV">
    ///     Type of the value retrieved from POCO. <see cref="DataType"/> is not the same if <typeparamref name="TV"/>
    ///     is <see cref="Nullable{T}"/>.
    /// </typeparam>
    public class PocoColumnSource<TA, TV> : IColumnDataSource
        where TA : class
    {
        /// <param name="memberDescriptor">
        ///     Mandatory; provides <see cref="ValueExtractor"/> and <see cref="Name"/>.
        /// </param>
        public PocoColumnSource(IExportedMemberDescriptor<TA, TV> memberDescriptor)
            : this (memberDescriptor.Name, memberDescriptor.ValueExtractor)
        {
            Check.DoCheckArgument(!memberDescriptor.IsCollection, () => $"Collections are not supported by this method ({memberDescriptor.Name}).");
        }

        /// <param name="name">
        ///     Optional
        /// </param>
        /// <param name="valueExtractor">
        ///     Mandatory, function extracting <typeparamref name="TV"/> from <typeparamref name="TA"/>.
        ///     The argument passed to the function will be null if actual POCO from which value is extracted
        ///     is not <typeparamref name="TA"/>.
        /// </param>
        public PocoColumnSource(string name, Func<TA, TV> valueExtractor)
        {
            Check.DoRequireArgumentNotNull(valueExtractor, nameof(valueExtractor));

            Name = name;
            ValueExtractor = valueExtractor;

            DataType = Nullable.GetUnderlyingType(typeof(TV)) ?? typeof(TV);
        }

        /// <inheritdoc />
        public string Name { get; set; }

        /// <inheritdoc />
        public Type DataType { get; }

        /// <inheritdoc />
        public object GetValue(object dataObject)
        {
            return ValueExtractor(dataObject as TA);
        }

        public Func<TA, TV> ValueExtractor { get; }
    }
}