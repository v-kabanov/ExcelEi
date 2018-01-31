// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-03-20
// Comment		
// **********************************************************************************************/

using System;
using System.Collections;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Defines column values source taking them from POCOs.
    /// </summary>
    public class PocoColumnSource : IColumnDataSource
    {
        public PocoColumnSource(string name, Type dataType, Func<object, object> valueExtractor)
        {
            Check.DoRequireArgumentNotNull(valueExtractor, nameof(valueExtractor));
            Check.DoRequireArgumentNotNull(dataType, nameof(dataType));

            Name = name;
            DataType = dataType;
            ValueExtractor = valueExtractor;
        }

        public string Name { get; set; }
        public Type DataType { get; }
        public Func<object, object> ValueExtractor { get; }

        public static PocoColumnSource Create<TE, TD>(Expression<Func<TE, TD>> getter)
        {
            Check.DoRequireArgumentNotNull(getter, nameof(getter));

            var compiledGetter = LambdaExpressionCache.Compile(getter);

            return new PocoColumnSource((getter.Body as MemberExpression)?.Member.Name, getter.ReturnType, e => compiledGetter((TE)e));
        }

        /// <summary>
        ///     Helper factory method, creates column source definition from type and member name using reflection.
        /// </summary>
        /// <typeparam name="TE">
        ///     Type of objects supplying value for the column
        /// </typeparam>
        /// <param name="memberName">
        ///     Property or field name containing value for the column
        /// </param>
        /// <returns>
        ///     Column source
        /// </returns>
        public static PocoColumnSource CreateReflection<TE>(string memberName)
        {
            return CreateReflection(typeof(TE), memberName);
        }

        /// <summary>
        ///     Helper factory method, creates column source definition from type and member name using reflection.
        /// </summary>
        /// <param name="entityType">
        ///     Type of objects supplying value for the column
        /// </param>
        /// <param name="memberName">
        ///     Property or field name containing value for the column
        /// </param>
        /// <returns>
        ///     Column source
        /// </returns>
        public static PocoColumnSource CreateReflection(Type entityType, string memberName)
        {
            Check.DoRequireArgumentNotNull(entityType, nameof(entityType));
            Check.DoRequireArgumentNotNull(memberName, nameof(memberName));

            var memberInfo = entityType.GetMember(memberName).FirstOrDefault(i => i is PropertyInfo || i is FieldInfo);

            Check.DoCheckArgument(memberInfo != null, () => $"Property or field {memberName} not found in {entityType.Name}.");

            Func<object, object> valueExtractor;
            Type dataType;

            var propertyInfo = memberInfo as PropertyInfo;
            if (propertyInfo != null)
            {
                Check.DoCheckArgument(propertyInfo.CanRead, () => $"Property {memberName} of {entityType.Name} is not readable");
                dataType = propertyInfo.PropertyType;
                valueExtractor = e => propertyInfo.GetValue(e);
            }
            else
            {
                var fieldInfo = (FieldInfo)memberInfo;
                Debug.Assert(fieldInfo != null, nameof(fieldInfo) + " != null");
                dataType = fieldInfo.FieldType;
                valueExtractor = e => fieldInfo.GetValue(e);
            }

            var isCollection = dataType.IsArray;
            if (!isCollection && dataType != typeof(string))
                isCollection = typeof(IEnumerable).IsAssignableFrom(dataType);

            Check.DoCheckArgument(!isCollection, () => $"Collections are not supported by this method ({memberName}).");

            dataType = Nullable.GetUnderlyingType(dataType) ?? dataType;

            return new PocoColumnSource(memberName, dataType, valueExtractor);
        }
    }
}