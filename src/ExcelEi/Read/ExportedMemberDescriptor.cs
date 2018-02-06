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
        public ExportedMemberDescriptor(Func<TA, TV> valueExtractor, string name)
        {
            Check.DoRequireArgumentNotNull(valueExtractor, nameof(valueExtractor));

            DataType = Nullable.GetUnderlyingType(typeof(TV)) ?? typeof(TV);
            ValueExtractor = valueExtractor;

            IsCollection = DataType.IsArray
                           || DataType != typeof(string) && typeof(IEnumerable).IsAssignableFrom(DataType);

            Name = name;
        }
    }
}