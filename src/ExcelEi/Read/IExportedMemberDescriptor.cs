// /**********************************************************************************************
// Author:  Vasily Kabanov
// Created  2018-02-06
// Comment  
// **********************************************************************************************/

using System;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Describes attributes of virtual column derived from POCO.
    /// </summary>
    /// <typeparam name="TA">
    ///     POCO type
    /// </typeparam>
    /// <typeparam name="TV">
    ///     Type of the value retrieved from POCO. <see cref="DataType"/> is not the same if <typeparamref name="TV"/>
    ///     is <see cref="Nullable{T}"/>.
    /// </typeparam>
    public interface IExportedMemberDescriptor<in TA, out TV>
        where TA : class
    {
        /// <summary>
        ///     Whether collection of values is extracted; string is not considered a collection.
        /// </summary>
        bool IsCollection { get; }

        /// <summary>
        ///     Extracted type, underlying for nullables.
        ///     E.g. if member is 'double?', typeof(double) will be returned.
        /// </summary>
        Type DataType { get; }

        Func<TA, TV> ValueExtractor { get; }

        /// <summary>
        ///     May be null for e.g. calculated members.
        /// </summary>
        string Name { get; }
    }
}