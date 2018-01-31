// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-03-20
// Comment		
// **********************************************************************************************/

using System;
using System.Collections.Concurrent;
using System.Linq.Expressions;

namespace ExcelEi
{
    /// <summary>
    ///     Caches compiled expressions; no housekeeping; dynamically constructed expressions should not be cached
    ///     in this implementation if their quantity is potentially unlimited.
    /// </summary>
    public static class LambdaExpressionCache
    {
        private static readonly ConcurrentDictionary<LambdaExpression, Delegate> Cache = new ConcurrentDictionary<LambdaExpression, Delegate>();

        private static Delegate CompileInternal(LambdaExpression expression)
        {
            Check.DoRequireArgumentNotNull(expression, nameof(expression));

            return Cache.GetOrAdd(expression, e => e.Compile());
        }

        /// <summary>
        ///     Return cached or compile and cache.
        /// </summary>
        /// <typeparam name="TA">
        ///     Expression's argument type
        /// </typeparam>
        /// <typeparam name="TR">
        ///     Expressions's result type
        /// </typeparam>
        /// <param name="expression">
        ///     Mandatory
        /// </param>
        /// <returns>
        ///     Compiled function
        /// </returns>
        public static Func<TA, TR> Compile<TA, TR>(Expression<Func<TA, TR>> expression)
        {
            Check.DoRequireArgumentNotNull(expression, nameof(expression));

            return (Func<TA, TR>)CompileInternal(expression);
        }
    }
}