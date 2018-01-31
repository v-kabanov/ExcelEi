// /**********************************************************************************************
// Author:  Vasily Kabanov
// Created  2017-04-27
// Comment  
// **********************************************************************************************/
// 

using System.Collections.Generic;

namespace ExcelEi
{
    public static class DictionaryExtensions
    {
        /// <summary>
        ///     Soft version of built-in method, returns default value if key not found.
        /// </summary>
        /// <param name="source">
        ///     Mandatory
        /// </param>
        /// <param name="key">
        ///     Key value, may be null
        /// </param>
        /// <returns>
        ///     default value for <typeparamref name="TValue"/> if <paramref name="key"/> is not contained in <paramref name="source"/>.
        /// </returns>
        public static TValue TryGetValue<TKey, TValue>(this IDictionary<TKey, TValue> source, TKey key)
        {
            Check.DoRequireArgumentNotNull(source, nameof(source));

            TValue result;
            if (source.TryGetValue(key, out result))
                return result;

            return default(TValue);
        }
    }
}