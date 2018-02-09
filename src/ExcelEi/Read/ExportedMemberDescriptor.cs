// /**********************************************************************************************
// Author:  Vasily Kabanov
// Created  2018-02-06
// Comment  
// **********************************************************************************************/

using System;
using System.Collections;

namespace ExcelEi.Read
{
    public class ExportedMemberDescriptor<TA, TV> : IExportedMemberDescriptor<TA, TV>
        where TA : class
    {
        /// <inheritdoc />
        public bool IsCollection { get; }

        /// <inheritdoc />
        public Type DataType { get; }

        /// <inheritdoc />
        public Func<TA, TV> ValueExtractor { get; }

        /// <inheritdoc />
        public string Name { get; }

        /// <inheritdoc />
        /// <param name="valueExtractor">
        ///     Mandatory
        /// </param>
        /// <param name="name">
        ///     Optional
        /// </param>
        /// <param name="dataType">
        ///     Optional, type of value extracted by <paramref name="valueExtractor"/>, if it is not
        ///     known at compile time.
        /// </param>
        public ExportedMemberDescriptor(Func<TA, TV> valueExtractor, string name, Type dataType = null)
        {
            Check.DoRequireArgumentNotNull(valueExtractor, nameof(valueExtractor));

            dataType = dataType ?? typeof(TV);

            DataType = Nullable.GetUnderlyingType(dataType) ?? dataType;
            ValueExtractor = valueExtractor;

            IsCollection = DataType.IsArray
                           || DataType != typeof(string) && typeof(IEnumerable).IsAssignableFrom(DataType);

            Name = name;
        }
    }
}