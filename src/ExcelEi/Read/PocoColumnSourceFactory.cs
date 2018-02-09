// /**********************************************************************************************
// Author:  Vasily Kabanov
// Created  2018-02-06
// Comment  
// **********************************************************************************************/

using System;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Class made for lessening verbosity by enabling type inference for generics.
    /// </summary>
    public static class PocoColumnSourceFactory
    {
        /// <summary>
        ///     Create column source object from expression extracting data from POCO. If <paramref name="getter"/> is
        ///     a simple reference to a property or field, <see cref="PocoColumnSource{TA,TV}.Name"/> will be set to the member name.
        ///     Otherwise it will be null.
        /// </summary>
        /// <typeparam name="TA">
        ///     POCO type; if an instance presented for extraction is not <typeparamref name="TA"/>, null is passed 
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Type of the value retrieved from POCO.
        /// </typeparam>
        /// <param name="getter">
        ///     Expression extracting an atomic value from POCO; could be a simple reference to a field or property
        ///     or calculation expression.
        /// </param>
        /// <param name="name">
        ///     Optional, name of the column in the virtual data table. By default it will be derived from referenced
        ///     member name.
        /// </param>
        /// <param name="preserveNameIfResolved">
        ///     Instructs that <paramref name="name"/> must not override resolved name, only be set if member name is
        ///     not resolved (prefer resolved name).
        /// </param>
        public static PocoColumnSource<TA, TV> Create<TA, TV>(
            Expression<Func<TA, TV>> getter, string name = null, bool preserveNameIfResolved = true)
            where TA : class
        {
            Check.DoRequireArgumentNotNull(getter, nameof(getter));

            var compiledGetter = LambdaExpressionCache.Compile(getter);

            var memberInfo = ExpressionHelper.GetMember(getter);

            var memberDescriptor = new ExportedMemberDescriptor<TA, TV>(compiledGetter, memberInfo?.Name);

            var result = new PocoColumnSource<TA, TV>(memberDescriptor);

            var ifNameResolved = !string.IsNullOrEmpty(result.Name);

            // if name is resolved and preserve option is ON, leave resolved name
            if (!string.IsNullOrEmpty(name) && (!ifNameResolved || !preserveNameIfResolved))
                result.Name = name;

            return result;
        }

        /// <summary>
        ///     Helper factory method, creates column source definition from type and member name using reflection.
        /// </summary>
        /// <typeparam name="TV">
        ///     Expected and extracted value type; if <see cref="object"/>, source data type will be set using reflection.
        ///     Otherwise actual value will be explicitly cast to <see cref="TV"/> when extracting.
        /// </typeparam>
        /// <param name="pocoType">
        ///     Type containing property or field.
        /// </param>
        /// <param name="memberName">
        ///     Property or field name containing value for the column
        /// </param>
        public static PocoColumnSource<object, TV> CreateReflection<TV>(Type pocoType, string memberName)
        {
            var memberDescriptor = CreateReflectionMemberDescriptor<TV>(pocoType, memberName);

            return new PocoColumnSource<object, TV>(memberDescriptor);
        }

        /// <summary>
        ///     Describe a property or field via reflection.
        /// </summary>
        /// <param name="pocoType">
        ///     Type containing property or field. If property it must not be indexed.
        /// </param>
        /// <param name="memberName">
        ///     Mandatory, name of existing property or field.
        /// </param>
        /// <exception cref="ArgumentException">
        ///     Member identified by <paramref name="memberName"/> cannot be read due to: <br />
        ///         - it does not exist <br />
        ///         - or it's not readable <br />
        ///         - or it is an indexed property <br />
        ///         - or it cannot be implicitly converted to <typeparamref name="TV"/> <br />
        /// </exception>
        public static IExportedMemberDescriptor<object, TV> CreateReflectionMemberDescriptor<TV>(Type pocoType, string memberName)
        {
            var memberInfo = pocoType.GetMember(memberName).FirstOrDefault(i => i is PropertyInfo || i is FieldInfo);

            Check.DoCheckArgument(memberInfo != null, () => $"Property or field {memberName} not found in {pocoType.Name}.");

            Func<object, TV> valueExtractor;
            Type dataType;

            var propertyInfo = memberInfo as PropertyInfo;
            if (propertyInfo != null)
            {
                Check.DoCheckArgument(propertyInfo.GetIndexParameters().Length == 0
                    , () => $"Property {memberName} of {pocoType.Name} is indexed which is not supported.");
                Check.DoCheckArgument(propertyInfo.CanRead, () => $"Property {memberName} of {pocoType.Name} is not readable");
                dataType = propertyInfo.PropertyType;
                valueExtractor = e => (TV) propertyInfo.GetValue(e);
            }
            else
            {
                var fieldInfo = memberInfo as FieldInfo;
                Check.DoCheckArgument(fieldInfo != null, () => $"Field or property {memberName} was not found on {pocoType.Name}.");
                Debug.Assert(fieldInfo != null, nameof(fieldInfo) + " != null");
                dataType = fieldInfo.FieldType;
                valueExtractor = e => (TV) fieldInfo.GetValue(e);
            }

            Debug.Assert(dataType != null);
            Check.DoCheckArgument(typeof(TV).IsAssignableFrom(dataType), () => $"Member {memberName} of {pocoType.Name} is " +
                                                                               // ReSharper disable once AccessToModifiedClosure
                                                                               $"{dataType.Name} and cannot be implicitly cast to {typeof(TV).Name}");

            var descriptorDataType = typeof(TV) == typeof(object)
                ? dataType
                : typeof(TV);

            return new ExportedMemberDescriptor<object, TV>(valueExtractor, memberName, descriptorDataType);
        }
    }
}