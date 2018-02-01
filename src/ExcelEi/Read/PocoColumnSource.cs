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
    public class PocoColumnSource<TA, TV> : IColumnDataSource
        where TA : class
    {
        public PocoColumnSource(string name, Func<TA, TV> valueExtractor)
        {
            Check.DoRequireArgumentNotNull(valueExtractor, nameof(valueExtractor));

            Name = name;
            ValueExtractor = valueExtractor;

            DataType = Nullable.GetUnderlyingType(typeof(TV)) ?? typeof(TV);
        }

        public string Name { get; set; }

        public Type DataType { get; }

        /// <inheritdoc />
        public object GetValue(object dataObject)
        {
            return ValueExtractor(dataObject as TA);
        }

        public Func<TA, TV> ValueExtractor { get; }

        public static PocoColumnSource<TA, TV> Create(Expression<Func<TA, TV>> getter)
        {
            Check.DoRequireArgumentNotNull(getter, nameof(getter));

            var compiledGetter = LambdaExpressionCache.Compile(getter);

            return new PocoColumnSource<TA, TV>((getter.Body as MemberExpression)?.Member.Name, compiledGetter);
        }

        /// <summary>
        ///     Helper factory method, creates column source definition from type and member name using reflection.
        /// </summary>
        /// <param name="memberName">
        ///     Property or field name containing value for the column
        /// </param>
        /// <returns>
        ///     Column source
        /// </returns>
        public static PocoColumnSource<TA, TV> CreateReflection(string memberName)
        {
            Check.DoRequireArgumentNotNull(memberName, nameof(memberName));

            var entityType = typeof(TA);

            var memberInfo = entityType.GetMember(memberName).FirstOrDefault(i => i is PropertyInfo || i is FieldInfo);

            Check.DoCheckArgument(memberInfo != null, () => $"Property or field {memberName} not found in {entityType.Name}.");

            Func<TA, TV> valueExtractor;
            Type dataType;

            var propertyInfo = memberInfo as PropertyInfo;
            if (propertyInfo != null)
            {
                Check.DoCheckArgument(propertyInfo.CanRead, () => $"Property {memberName} of {entityType.Name} is not readable");
                dataType = propertyInfo.PropertyType;
                valueExtractor = e => (TV)propertyInfo.GetValue(e);
            }
            else
            {
                var fieldInfo = (FieldInfo)memberInfo;
                Debug.Assert(fieldInfo != null, nameof(fieldInfo) + " != null");
                dataType = fieldInfo.FieldType;
                valueExtractor = e => (TV)fieldInfo.GetValue(e);
            }

            Debug.Assert(dataType != null);
            Check.DoCheckArgument(typeof(TV).IsAssignableFrom(dataType), () => $"Member {memberName} of {entityType.Name} is " +
                                                                               // ReSharper disable once AccessToModifiedClosure
                                                                               $"{dataType.Name} and cannot be implicitly cast to {typeof(TV).Name}");

            var isCollection = dataType.IsArray;
            if (!isCollection && dataType != typeof(string))
                isCollection = typeof(IEnumerable).IsAssignableFrom(dataType);

            Check.DoCheckArgument(!isCollection, () => $"Collections are not supported by this method ({memberName}).");

            dataType = Nullable.GetUnderlyingType(dataType) ?? dataType;

            return new PocoColumnSource<TA, TV>(memberName, valueExtractor);
        }
    }
}