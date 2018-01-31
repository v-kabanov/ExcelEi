// /**********************************************************************************************
// Author:  Vasily Kabanov
// Created  2018-01-10
// Comment  
// **********************************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelEi.Read;

namespace ExcelEi.Write
{
    /// <summary>
    ///     Stateful object allowing to build configuration for export from POCO collection to
    ///     Excel incrementally.
    /// </summary>
    public class PocoExportConfigurator
    {
        public const int MaxSheetNameLength = 31;

        /// <summary>
        ///     Use sheet name as source table name.
        /// </summary>
        /// <param name="pocoType">
        ///     Mandatory
        /// </param>
        /// <param name="sheetName">
        ///     Excel sheet and virtual source table name
        /// </param>
        public PocoExportConfigurator(Type pocoType, string sheetName)
            : this(pocoType, sheetName, sheetName)
        {
        }

        /// <summary>
        ///     Use different sheet and data table names.
        /// </summary>
        /// <param name="pocoType">
        ///     Mandatory
        /// </param>
        /// <param name="sheetName">
        ///     Excel sheet and virtual source table name
        /// </param>
        /// <param name="dataTableName">
        ///     Name of the table in source data set (key in <see cref="IDataSet.DataTables"/>), see also <see cref="DataSetAdapter.Add(IDataTable,string)"/>.
        /// </param>
        public PocoExportConfigurator(Type pocoType, string sheetName, string dataTableName)
        {
            Check.DoRequireArgumentNotNull(pocoType, nameof(pocoType));
            Check.DoRequireArgumentNotBlank(sheetName, nameof(sheetName));
            Check.DoRequireArgumentNotBlank(dataTableName, nameof(dataTableName));
            Check.DoCheckArgument(sheetName.Length <= MaxSheetNameLength
                , () => $"Sheet name must not exceed {MaxSheetNameLength} characters in length.");
            
            PocoType = pocoType;

            Config = new DataTableExportAutoConfig
            {
                SheetName = sheetName,
                DataTableName = dataTableName
            };
        }

        public Type PocoType { get; }

        /// <summary>
        ///     Configuration being constructed.
        /// </summary>
        public DataTableExportAutoConfig Config { get; }

        /// <summary>
        ///     Add custom column sourcing data for it from lambda expression.
        /// </summary>
        /// <typeparam name="TV"></typeparam>
        /// <param name="valueExtractor">
        ///     Mandatory, extracts data from data source objects.
        /// </param>
        /// <param name="sheetColumnIndex">
        ///     0-based sheet column index relative to left most column to which export is performed
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel. Also goes into <see cref="IColumnDataSource"/>'s <see cref="IColumnDataSource.Name"/>.
        /// </param>
        /// <param name="autoFit">
        ///     Whether to fit column width to content after export.
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns></returns>
        public DataColumnExportAutoConfig AddColumn<TV>(Func<object, TV> valueExtractor, int sheetColumnIndex, string sheetColumnCaption, bool? autoFit, string format)
        {
            var columnSource = new PocoColumnSource(sheetColumnCaption, typeof(TV), o => valueExtractor(o));
            var config = Add(columnSource, sheetColumnIndex, sheetColumnCaption, autoFit, format);

            return config;
        }

        /// <summary>
        ///     Add column getting values from property of non-collection type by reflection
        /// </summary>
        /// <param name="memberName">
        ///     Field or property name; must be of primitive type, not a colection.
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel.
        /// </param>
        /// <param name="autoFit">
        ///     Whether to fit column width to content after export.
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns></returns>
        public PocoExportConfigurator AddColumn(string memberName, string sheetColumnCaption = null, bool? autoFit = null, string format = null)
        {
            Check.DoRequireArgumentNotNull(memberName, nameof(memberName));

            var columnSource = PocoColumnSource.CreateReflection(PocoType, memberName);

            Add(columnSource, Config.Columns.Count, sheetColumnCaption, autoFit, format);

            return this;
        }

        /// <summary>
        ///     Add column getting values from property of non-collection type by lambda expression which
        ///     is compiled and cached.
        /// </summary>
        /// <param name="getter">
        ///     Mandatory, reference returning property or field or arbitrary value. In the latter case
        ///     <paramref name="sheetColumnCaption"/> should be specified and it will be used as source column name
        ///     for identification. The expression is compiled, cached and used for retrieving values for the column.
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel.
        /// </param>
        /// <param name="autoFit">
        ///     Whether to fit column width to content after export.
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        public PocoExportConfigurator AddColumn<TA, TR>(Expression<Func<TA, TR>> getter, string sheetColumnCaption = null, bool? autoFit = null, string format = null)
        {
            Check.DoRequireArgumentNotNull(getter, nameof(getter));

            var columnSource = PocoColumnSource.Create(getter);
            if (string.IsNullOrEmpty(columnSource.Name))
            {
                columnSource.Name = sheetColumnCaption;
            }

            Add(columnSource, Config.Columns.Count, sheetColumnCaption, autoFit, format);

            return this;
        }

        /// <summary>
        ///     Configure export of collection member with up to <paramref name="columnCount"/> elements.
        ///     A column will be created for every collection item.
        /// </summary>
        /// <param name="memberName">
        ///     Name of property or field returning array or <see cref="IList"/>, mandatory
        /// </param>
        /// <param name="columnCount">
        ///     Number of columns to create. Collection elements exceeding this limit will not be exported.
        /// </param>
        /// <param name="sheetColumnCaptionBase">
        ///     Optional, default is same as <paramref name="memberName"/>; columns will be named by appending '[index]' to the base.
        /// </param>
        /// <param name="autoFit"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public PocoExportConfigurator AddCollectionColumns(string memberName, int columnCount, string sheetColumnCaptionBase, bool? autoFit = null, string format = null)
        {
            Check.DoRequireArgumentNotBlank(memberName, nameof(memberName));
            Check.DoCheckArgument(columnCount > 0, nameof(columnCount));

            sheetColumnCaptionBase = sheetColumnCaptionBase ?? memberName;

            var memberInfo = PocoType.GetMember(memberName).FirstOrDefault(i => i is PropertyInfo || i is FieldInfo);
            var propertyInfo = memberInfo as PropertyInfo;
            Type collectionType;
            Func<object, IList> collectionExtractor;
            if (propertyInfo != null)
            {
                Check.DoCheckArgument(propertyInfo.CanRead, () => $"Property {memberName} of {PocoType.Name} is not readable");
                collectionType = propertyInfo.PropertyType;
                collectionExtractor = e => (IList)propertyInfo.GetValue(e);
            }
            else
            {
                var fieldInfo = (FieldInfo)memberInfo;
                Debug.Assert(fieldInfo != null, nameof(fieldInfo) + " != null");
                collectionType = fieldInfo.FieldType;
                collectionExtractor = e => (IList)fieldInfo.GetValue(e);
            }

            Check.DoCheckArgument(collectionType.IsArray || typeof(IList).IsAssignableFrom(collectionType)
                , () => $"{PocoType.Name}.{memberName} is not a collection.");

            var elementType = GetCollectionElementType(collectionType);

            elementType = Nullable.GetUnderlyingType(elementType) ?? elementType;

            AddCollectionColumns(collectionExtractor, elementType, columnCount, sheetColumnCaptionBase, autoFit, format);

            return this;
        }

        /// <summary>
        ///     Configure export of collection member with up to <paramref name="columnCount"/> elements.
        ///     A column will be created for every collection item, limited by max number supported by excel.
        ///     This version supports arrays and lists (both generic and non-generic).
        /// </summary>
        /// <param name="collectionMemberGetter">
        ///     Name of property or field returning array or <see cref="IList"/>, mandatory. If 
        ///     Mandatory, reference returning property or field or arbitrary array or <see cref="IList"/> typed value.
        ///     If it is not a property or field reference, <paramref name="sheetColumnCaptionBase"/> must be specified
        ///     and it will be used as source column name base for identification.
        ///     The expression is compiled, cached and used for retrieving values for the column.
        /// </param>
        /// <param name="columnCount">
        ///     Number of columns to create. Collection elements exceeding this limit will not be exported.
        ///     Max number of columns supported by excel is 16384.
        /// </param>
        /// <param name="sheetColumnCaptionBase">
        ///     Optional, default is same as <paramref name="collectionMemberGetter"/>'s name; columns will be named by appending '[index]' to the base.
        /// </param>
        /// <param name="autoFit"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public PocoExportConfigurator AddCollectionColumns<TA>(Expression<Func<TA, IList>> collectionMemberGetter, int columnCount, string sheetColumnCaptionBase = null, bool? autoFit = null, string format = null)
        {
            Check.DoRequireArgumentNotNull(collectionMemberGetter, nameof(collectionMemberGetter));
            Check.DoCheckArgument(columnCount > 0 && columnCount < 16384, nameof(columnCount));
            Check.DoCheckArgument(collectionMemberGetter.Body is MemberExpression || !string.IsNullOrWhiteSpace(sheetColumnCaptionBase)
                                  , "When using non-member expression name must be provided");

            var memberInfo = (collectionMemberGetter.Body as MemberExpression)?.Member
                             ?? ((collectionMemberGetter.Body as UnaryExpression)?.Operand as MemberExpression)?.Member;

            Check.DoCheckArgument(memberInfo != null || !string.IsNullOrWhiteSpace(sheetColumnCaptionBase)
                                  , () => "If collection expression does not refer to property or field, sheet column caption base must be provided");

            if (string.IsNullOrWhiteSpace(sheetColumnCaptionBase))
            {
                Debug.Assert(memberInfo != null, nameof(memberInfo) + " != null");
                sheetColumnCaptionBase = memberInfo.Name;
            }

            var memberName = memberInfo?.Name ?? sheetColumnCaptionBase;

            var propertyInfo = memberInfo as PropertyInfo;
            Type collectionType;
            Func<object, IList> collectionExtractor;
            if (propertyInfo != null)
            {
                Check.DoCheckArgument(propertyInfo.CanRead, () => $"Property {memberName} of {PocoType.Name} is not readable");
                collectionType = propertyInfo.PropertyType;
                collectionExtractor = e => (IList)propertyInfo.GetValue(e);
            }
            else
            {
                var fieldInfo = (FieldInfo)memberInfo;
                Debug.Assert(fieldInfo != null, nameof(fieldInfo) + " != null");
                collectionType = fieldInfo.FieldType;
                collectionExtractor = e => (IList)fieldInfo.GetValue(e);
            }

            Check.DoCheckArgument(collectionType.IsArray || typeof(IList).IsAssignableFrom(collectionType)
                , () => $"{PocoType.Name}.{memberName} is not a collection.");

            var elementType = GetCollectionElementType(collectionType);

            elementType = Nullable.GetUnderlyingType(elementType) ?? elementType;

            AddCollectionColumns(collectionExtractor, elementType, columnCount, sheetColumnCaptionBase, autoFit, format);

            return this;
        }

        /// <summary>
        ///     Configure export of generic list (or array) member with up to <paramref name="columnCount"/> elements.
        ///     A column will be created for every collection item, limited by max number supported by excel.
        /// </summary>
        /// <param name="collectionMemberGetter">
        ///     Name of property or field returning array or <see cref="IList{TV}"/>, mandatory. If 
        ///     Mandatory, reference returning property or field implementing <see cref="IList{TV}"/>.
        ///     If it is not a property or field reference, <paramref name="sheetColumnCaptionBase"/> must be specified
        ///     and it will be used as source column name base for identification.
        ///     The expression is compiled, cached and used for retrieving values for the column.
        /// </param>
        /// <param name="columnCount">
        ///     Number of columns to create. Collection elements exceeding this limit will not be exported.
        ///     Max number of columns supported by excel is 16384.
        /// </param>
        /// <param name="sheetColumnCaptionBase">
        ///     Optional, default is same as <paramref name="collectionMemberGetter"/>'s name; columns will be named by appending '[index]' to the base.
        /// </param>
        /// <param name="autoFit"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public PocoExportConfigurator AddCollectionColumns<TA, TV>(Expression<Func<TA, IList<TV>>> collectionMemberGetter, int columnCount, string sheetColumnCaptionBase = null, bool? autoFit = null, string format = null)
        {
            Check.DoRequireArgumentNotNull(collectionMemberGetter, nameof(collectionMemberGetter));
            Check.DoCheckArgument(columnCount > 0 && columnCount < 16384, nameof(columnCount));
            Check.DoCheckArgument(collectionMemberGetter.Body is MemberExpression || !string.IsNullOrWhiteSpace(sheetColumnCaptionBase)
                                  , "When using non-member expression name must be provided");

            var memberInfo = (collectionMemberGetter.Body as MemberExpression)?.Member
                             ?? ((collectionMemberGetter.Body as UnaryExpression)?.Operand as MemberExpression)?.Member;

            Check.DoCheckArgument(memberInfo != null || !string.IsNullOrWhiteSpace(sheetColumnCaptionBase)
                                  , () => "If collection expression does not refer to property or field, sheet column caption base must be provided");

            if (string.IsNullOrWhiteSpace(sheetColumnCaptionBase))
            {
                Debug.Assert(memberInfo != null, nameof(memberInfo) + " != null");
                sheetColumnCaptionBase = memberInfo.Name;
            }

            var memberName = memberInfo?.Name ?? sheetColumnCaptionBase;

            var propertyInfo = memberInfo as PropertyInfo;
            Type collectionType;
            Func<object, IList<TV>> collectionExtractor;
            if (propertyInfo != null)
            {
                Check.DoCheckArgument(propertyInfo.CanRead, () => $"Property {memberName} of {PocoType.Name} is not readable");
                collectionType = propertyInfo.PropertyType;
                collectionExtractor = e => (IList<TV>)propertyInfo.GetValue(e);
            }
            else
            {
                var fieldInfo = (FieldInfo)memberInfo;
                Debug.Assert(fieldInfo != null, nameof(fieldInfo) + " != null");
                collectionType = fieldInfo.FieldType;
                collectionExtractor = e => (IList<TV>)fieldInfo.GetValue(e);
            }

            Check.DoCheckArgument(typeof(IList<TV>).IsAssignableFrom(collectionType)
                , () => $"{PocoType.Name}.{memberName} is not a generic list.");

            var elementType = GetCollectionElementType(collectionType);

            elementType = Nullable.GetUnderlyingType(elementType) ?? elementType;

            AddCollectionColumns(collectionExtractor, elementType, columnCount, sheetColumnCaptionBase, autoFit, format);

            return this;
        }

        /// <summary>
        ///     Low level method allowing client to configure column source in any way they need.
        /// </summary>
        /// <param name="columnDataSource">
        ///     Encapsulates data retrieval for the column.
        /// </param>
        /// <param name="sheetColumnIndex">
        ///     0-based sheet column index relative to left most column to which export is performed
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel.
        /// </param>
        /// <param name="autoFit">
        ///     Whether to fit column width to content after export.
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        public DataColumnExportAutoConfig Add(IColumnDataSource columnDataSource, int sheetColumnIndex, string sheetColumnCaption, bool? autoFit, string format)
        {
            Check.DoRequireArgumentNotNull(columnDataSource, nameof(columnDataSource));

            if (string.IsNullOrEmpty(sheetColumnCaption) && !string.IsNullOrEmpty(columnDataSource.Name))
                sheetColumnCaption = columnDataSource.Name;

            var config = new DataColumnExportAutoConfig(Config, sheetColumnIndex, sheetColumnCaption, columnDataSource);

            if (autoFit.HasValue)
                config.AutoFit = autoFit.Value;

            if (format != null)
                config.Format = format;

            Config.AddColumn(config);

            return config;
        }

        private TV TryGetCollectionElement<TV>(IList<TV> list, int index)
        {
            if (list == null || index >= list.Count || index < 0)
                return default(TV);

            return list[index];
        }

        private object TryGetCollectionElement(IList list, int index)
        {
            if (list == null || index >= list.Count || index < 0)
                return null;

            return list[index];
        }

        private Type GetCollectionElementType(Type collectionType)
        {
            Check.DoRequireArgumentNotNull(collectionType, nameof(collectionType));

            var collectionIfaceType = collectionType.GetGenericTypeDefinition() == typeof(IList<>)
                                        ? collectionType
                                        : collectionType.GetInterfaces().FirstOrDefault(
                                            i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IList<>));

            Check.DoCheckArgument(typeof(IList).IsAssignableFrom(collectionType)
                                  || collectionIfaceType != null, () => $"Type {collectionType.FullName} doesn not implement IList or IList<>.");

            if (collectionType.IsArray)
                return collectionType.GetElementType();

            return collectionIfaceType?.GetGenericArguments()[0] ?? typeof(object);
        }

        private void AddCollectionColumns<TV>(Func<object, IList<TV>> collectionGetter,
            Type elementType,
            int columnCount,
            string sheetColumnCaptionBase,
            bool? autoFit = null,
            string format = null)
        {
            for (var i = 0; i < columnCount; ++i)
            {
                var collectionIndex = i;
                var columnName = $"{sheetColumnCaptionBase}[{collectionIndex}]";
                var columnSource = new PocoColumnSource(columnName, elementType, o => TryGetCollectionElement(collectionGetter(o), collectionIndex));
                Add(columnSource, Config.Columns.Count, columnName, autoFit, format);
            }
        }

        private void AddCollectionColumns(Func<object, IList> collectionGetter,
            Type elementType,
            int columnCount,
            string sheetColumnCaptionBase,
            bool? autoFit = null,
            string format = null)
        {
            for (var i = 0; i < columnCount; ++i)
            {
                var collectionIndex = i;
                var columnName = $"{sheetColumnCaptionBase}[{collectionIndex}]";
                var columnSource = new PocoColumnSource(columnName, elementType, o => TryGetCollectionElement(collectionGetter(o), collectionIndex));
                Add(columnSource, Config.Columns.Count, columnName, autoFit, format);
            }
        }
    }
}