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
    ///     is <see cref="Nullable{T}"/> or when using reflection. When reflection is used <typeparamref name="TV"/>
    ///     will often be <see cref="object"/> but <see cref="DataType"/> should be correctly reflected.
    /// </typeparam>
    public class PocoColumnSource<TA, TV> : IColumnDataSource
        where TA : class
    {
        /// <param name="memberDescriptor">
        ///     Mandatory; provides <see cref="ValueExtractor"/> and <see cref="Name"/>.
        /// </param>
        public PocoColumnSource(IExportedMemberDescriptor<TA, TV> memberDescriptor)
        {
            Check.DoCheckArgument(!memberDescriptor.IsCollection, () => $"Collections are not supported by this method ({memberDescriptor.Name}).");

            Name = memberDescriptor.Name;
            ValueExtractor = memberDescriptor.ValueExtractor;
            DataType = memberDescriptor.DataType;
        }

        /// <param name="name">
        ///     Optional
        /// </param>
        /// <param name="valueExtractor">
        ///     Mandatory, function extracting <typeparamref name="TV"/> from <typeparamref name="TA"/>.
        ///     The argument passed to the function will be null if actual POCO from which value is extracted
        ///     is not <typeparamref name="TA"/>.
        /// </param>
        /// <param name="dataType">
        ///     Optional, type of value extracted by <paramref name="valueExtractor"/>, if it is not
        ///     known at compile time.
        /// </param>
        public PocoColumnSource(string name, Func<TA, TV> valueExtractor, Type dataType = null)
        {
            Check.DoRequireArgumentNotNull(valueExtractor, nameof(valueExtractor));

            Name = name;
            ValueExtractor = valueExtractor;

            dataType = dataType ?? typeof(TV);

            DataType = Nullable.GetUnderlyingType(dataType) ?? dataType;
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