// /**********************************************************************************************
// Author:  Vasily Kabanov
// Created  2018-01-10
// Comment  
// **********************************************************************************************/

using System;
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
    ///     Excel incrementally. Note that if during export actual POCO type cannot be cast to the configured type
    ///     <see cref="PocoColumnSource{TA,TV}"/> will pass null to expression based value getters (simple references will produce
    ///     <see cref="NullReferenceException"/>) and reflection based ones will throw <see cref="TargetException"/> unless
    ///     the member is static.
    /// </summary>
    /// <typeparam name="TE">
    ///     This is expected type of POCOs.
    ///     Maximum type safety can be achieved when mapping from this or its ancestors.
    ///     It is possible to map data from descendant types but type safety will not be enforced at compile time.
    /// </typeparam>
    public class PocoExportConfigurator<TE>
        where TE: class
    {
        /// <summary>
        ///     Limit imposed by excel on the length of sheet name.
        /// </summary>
        public const int MaxSheetNameLength = 31;

        /// <summary>
        ///     Use sheet name as source table name.
        /// </summary>
        /// <param name="sheetName">
        ///     Excel sheet and virtual source table name
        /// </param>
        public PocoExportConfigurator(string sheetName)
            : this(sheetName, sheetName)
        {
        }

        /// <summary>
        ///     Use different sheet and data table names.
        /// </summary>
        /// <param name="sheetName">
        ///     Excel sheet and virtual source table name
        /// </param>
        /// <param name="dataTableName">
        ///     Name of the table in source data set (key in <see cref="IDataSet.DataTables"/>), see also <see cref="DataSetAdapter.Add(IDataTable,string)"/>.
        /// </param>
        public PocoExportConfigurator(string sheetName, string dataTableName)
        {
            Check.DoRequireArgumentNotBlank(sheetName, nameof(sheetName));
            Check.DoRequireArgumentNotBlank(dataTableName, nameof(dataTableName));
            Check.DoCheckArgument(sheetName.Length <= MaxSheetNameLength
                , () => $"Sheet name must not exceed {MaxSheetNameLength} characters in length.");
            
            PocoType = typeof(TE);

            Config = new DataTableExportAutoConfig
            {
                SheetName = sheetName,
                DataTableName = dataTableName
            };
        }

        /// <summary>
        ///     Type instance of <typeparamref name="TE"/>. This is expected type of POCOs.
        ///     Maximum type safety can be achieved when mapping from this or its ancestors.
        ///     It is possible to map data from descendant types but type safety will not be enforced at compile time.
        /// </summary>
        public Type PocoType { get; }

        /// <summary>
        ///     Configuration being constructed.
        /// </summary>
        public DataTableExportAutoConfig Config { get; }

        /// <summary>
        ///     Add custom column sourcing data for it from lambda expression.
        ///     Use to produce e.g. sparse columns by forcing <paramref name="sheetColumnIndex"/>.
        /// </summary>
        /// <typeparam name="TV">
        ///     Source data type - type returned by <paramref name="valueGetter"/>.
        /// </typeparam>
        /// <param name="valueGetter">
        ///     Mandatory, extracts data from data source objects.
        /// </param>
        /// <param name="sheetColumnIndex">
        ///     0-based sheet column index relative to left most column to which export is performed
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel. Also goes into <see cref="IColumnDataSource"/>'s <see cref="IColumnDataSource.Name"/>.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     New column configuration (can be further customized).
        /// </returns>
        public DataColumnExportAutoConfig AddColumn<TV>(Func<TE, TV> valueGetter, int sheetColumnIndex, string sheetColumnCaption, bool? autoFit, string format)
        {
            return AddColumnPolymorphic(valueGetter, sheetColumnIndex, sheetColumnCaption, autoFit, format);
        }

        /// <summary>
        ///     Add custom column sourcing data for it from lambda expression.
        ///     Use to produce e.g. sparse columns by forcing <paramref name="sheetColumnIndex"/>.
        /// </summary>
        /// <typeparam name="TA">
        ///     Instance type accepted by <paramref name="valueGetter"/>; must be same as <typeparamref name="TE"/> or its base.
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Source data type - type returned by <paramref name="valueGetter"/>.
        /// </typeparam>
        /// <param name="valueGetter">
        ///     Mandatory, extracts data from data source objects.
        /// </param>
        /// <param name="sheetColumnIndex">
        ///     0-based sheet column index relative to left most column to which export is performed
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel. Also goes into <see cref="IColumnDataSource"/>'s <see cref="IColumnDataSource.Name"/>.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     New column configuration (can be further customized).
        /// </returns>
        /// <remarks>
        ///     There should be really reverse type constraint 'TE: TA', TA should be base class for TE, but
        ///     I cannot find a way to express it.
        /// </remarks>
        public DataColumnExportAutoConfig AddInheritedColumn<TA, TV>(
            Func<TA, TV> valueGetter, int sheetColumnIndex, string sheetColumnCaption, bool? autoFit, string format)
            where TA: class
        {
            Check.DoCheckArgument(typeof(TA).IsAssignableFrom(PocoType), () => $"{PocoType.Name} does not inherit from {typeof(TA).Name}");

            return AddColumnPolymorphic(valueGetter, sheetColumnIndex, sheetColumnCaption, autoFit, format);
        }

        /// <summary>
        ///     Add custom column sourcing data for it from lambda expression.
        ///     Use to produce e.g. sparse columns by forcing <paramref name="sheetColumnIndex"/>.
        ///     The expression can accept <see cref="PocoType"/>'s descendant class instance which is checked at compile time;
        ///     however, whether exported collection item is actually of the descendant type
        ///     will only be checked at runtime during export (rather than configuration). If it cannot be cast to <typeparamref name="TA"/>,
        ///     null will be passed to the <paramref name="valueGetter"/> by <see cref="PocoColumnSource{TA,TV}"/>.
        /// </summary>
        /// <typeparam name="TA">
        ///     Instance type accepted by <paramref name="valueGetter"/>; must be same as <typeparamref name="TE"/> or its descendant.
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Source data type - type returned by <paramref name="valueGetter"/>.
        /// </typeparam>
        /// <param name="valueGetter">
        ///     Mandatory, extracts data from data source objects.
        /// </param>
        /// <param name="sheetColumnIndex">
        ///     0-based sheet column index relative to left most column to which export is performed
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel. Also goes into <see cref="IColumnDataSource"/>'s <see cref="IColumnDataSource.Name"/>.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     New column configuration (can be further customized).
        /// </returns>
        /// <remarks>
        ///     This method encapsulates generic type constraint.
        /// </remarks>
        public DataColumnExportAutoConfig AddDescendantsColumn<TA, TV>(
            Func<TA, TV> valueGetter, int sheetColumnIndex, string sheetColumnCaption, bool? autoFit, string format)
            where TA: class, TE
        {
            return AddColumnPolymorphic(valueGetter, sheetColumnIndex, sheetColumnCaption, autoFit, format);
        }

        /// <summary>
        ///     Add custom column sourcing data for it from lambda expression.
        ///     Use to produce e.g. sparse columns by forcing <paramref name="sheetColumnIndex"/>.
        /// </summary>
        /// <typeparam name="TA">
        ///     Instance type accepted by <paramref name="valueGetter"/>; must be same as <typeparamref name="TE"/>, its base or descendant.
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Source data type - type returned by <paramref name="valueGetter"/>.
        /// </typeparam>
        /// <param name="valueGetter">
        ///     Mandatory, extracts data from data source objects.
        /// </param>
        /// <param name="sheetColumnIndex">
        ///     0-based sheet column index relative to left most column to which export is performed
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel. Also goes into <see cref="IColumnDataSource"/>'s <see cref="IColumnDataSource.Name"/>.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     New column configuration (can be further customized).
        /// </returns>
        public DataColumnExportAutoConfig AddColumnPolymorphic<TA, TV>(
            Func<TA, TV> valueGetter, int sheetColumnIndex, string sheetColumnCaption, bool? autoFit, string format)
            where TA: class
        {
            Check.DoCheckArgument(typeof(TA).IsAssignableFrom(PocoType) || PocoType.IsAssignableFrom(typeof(TA))
                , () => $"{PocoType.Name} cannot be cast to {typeof(TA).Name}");

            var columnSource = new PocoColumnSource<TA, TV>(sheetColumnCaption, valueGetter);
            var config = AddColumn(columnSource, sheetColumnIndex, sheetColumnCaption, autoFit, format);

            return config;
        }

        /// <summary>
        ///     Add column getting values from property or field of non-collection type by reflection.
        ///     If actual POCO instance is not of <see cref="PocoType"/> type, <see cref="TargetException"/> will be thrown during export.
        ///     This is different from the behavior of methods accepting expressions because expressions can potentially handle null argument.
        /// </summary>
        /// <typeparam name="TV">
        ///     Source data type - type to which member identified by <see cref="memberName"/> will be cast.
        ///     Specify <see cref="object"/> to let it (<see cref="PocoColumnSource{TA,TV}.DataType"/>) be figured out via reflection;
        ///     actual value will not be cast during export.
        /// </typeparam>
        /// <param name="memberName">
        ///     Field or property name; must be of primitive type, not a colection.
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        public PocoExportConfigurator<TE> AddColumn<TV>(string memberName, string sheetColumnCaption = null, bool? autoFit = null, string format = null)
        {
            return AddDescendantsColumn<TE, TV>(memberName, sheetColumnCaption, autoFit, format);
        }

        /// <summary>
        ///     Add column getting values from a non-collection member of <see cref="PocoType"/>'s descendant type by reflection.
        ///     If actual POCO instance is not of the descendant type, <see cref="TargetException"/> will be thrown during export.
        ///     This is different from the behavior of methods accepting expressions because expressions can potentially handle null argument.
        /// </summary>
        /// <typeparam name="TA">
        ///     Descendant type containing <paramref name="memberName"/>.
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Source data type - type to which member identified by <see cref="memberName"/> will be cast.
        ///     Specify <see cref="object"/> to let it (<see cref="PocoColumnSource{TA,TV}.DataType"/>) be figured out via reflection;
        ///     actual value will not be cast during export.
        /// </typeparam>
        /// <param name="memberName">
        ///     Field or property name; must be of primitive type, not a colection.
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        public PocoExportConfigurator<TE> AddDescendantsColumn<TA, TV>(string memberName, string sheetColumnCaption = null, bool? autoFit = null, string format = null)
            where TA: class, TE
        {
            Check.DoRequireArgumentNotNull(memberName, nameof(memberName));

            var columnSource = PocoColumnSourceFactory.CreateReflection<TV>(typeof(TA), memberName);

            AddColumn(columnSource, Config.Columns.Count, sheetColumnCaption, autoFit, format);

            return this;
        }

        /// <summary>
        ///     Add column getting values from property of non-collection type by lambda expression which
        ///     is compiled and cached.
        /// </summary>
        /// <param name="valueGetter">
        ///     Mandatory, reference returning property (value conversions do not prevent member resolution) or field or arbitrary value.
        ///     In the latter case <paramref name="sheetColumnCaption"/> should be specified and it will be used as source column name
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
        public PocoExportConfigurator<TE> AddColumn<TV>(Expression<Func<TE, TV>> valueGetter, string sheetColumnCaption = null, bool? autoFit = null, string format = null)
        {
            return AddInheritedColumn(valueGetter, sheetColumnCaption, autoFit, format);
        }

        /// <summary>
        ///     Add column getting values from property of non-collection type by lambda expression which
        ///     is compiled and cached. The expression can accept <see cref="PocoType"/>'s base class instance; in this
        ///     case type safety will be checked at runtime during configuration as it cannot be enforced at compile time.
        /// </summary>
        /// <typeparam name="TA">
        ///     Instance type accepted by <paramref name="valueGetter"/>; same as <typeparamref name="TE"/> or its base class
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Source data type - type returned by <paramref name="valueGetter"/>.
        /// </typeparam>
        /// <param name="valueGetter">
        ///     Mandatory, reference returning property (value conversions do not prevent member resolution) or field or arbitrary value.
        ///     In the latter case <paramref name="sheetColumnCaption"/> should be specified and it will be used as source column name
        ///     for identification. The expression is compiled, cached and used for retrieving values for the column.
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        /// <remarks>
        ///     There should be really reverse type constraint 'TE: TA', TA should be base class for TE, but
        ///     I cannot find a way to express it.
        /// </remarks>
        public PocoExportConfigurator<TE> AddInheritedColumn<TA, TV>(Expression<Func<TA, TV>> valueGetter, string sheetColumnCaption = null, bool? autoFit = null, string format = null)
            where   TA: class
        {
            Check.DoRequireArgumentNotNull(valueGetter, nameof(valueGetter));
            Check.DoCheckArgument(typeof(TA).IsAssignableFrom(PocoType), () => $"{PocoType.Name} does not inherit from {typeof(TA).Name}");

            var columnSource = PocoColumnSourceFactory.Create(valueGetter, sheetColumnCaption);

            return AddColumn(columnSource, sheetColumnCaption, autoFit, format);
        }

        /// <summary>
        ///     Add column getting values from property of non-collection type by lambda expression which
        ///     is compiled and cached. The expression can accept <see cref="PocoType"/>'s descendant class instance
        ///     which is checked at compile time; however, whether exported collection item is actually of the descendant type
        ///     will only be checked at runtime during export (rather than configuration). If it cannot be cast to <typeparamref name="TA"/>,
        ///     null will be passed to the <paramref name="valueGetter"/> by <see cref="PocoColumnSource{TA,TV}"/>.
        /// </summary>
        /// <typeparam name="TA">
        ///     Instance type accepted by <paramref name="valueGetter"/>; must be same as <typeparamref name="TE"/> or its descendant.
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Source data type - type returned by <paramref name="valueGetter"/>.
        /// </typeparam>
        /// <param name="valueGetter">
        ///     Mandatory, reference returning property (value conversions do not prevent member resolution) or field or arbitrary value.
        ///     In the latter case <paramref name="sheetColumnCaption"/> should be specified and it will be used as source column name
        ///     for identification. The expression is compiled, cached and used for retrieving values for the column.
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        /// <remarks>
        ///     This method encapsulates generic type constraint.
        /// </remarks>
        public PocoExportConfigurator<TE> AddDescendantsColumn<TA, TV>(
            Expression<Func<TA, TV>> valueGetter
            , string sheetColumnCaption = null, bool? autoFit = null, string format = null)
            where   TA: class, TE
        {
            return AddColumnPolymorphic(valueGetter, sheetColumnCaption, autoFit, format);
        }

        /// <summary>
        ///     Add column getting values from property of non-collection type by lambda expression which
        ///     is compiled and cached. The expression can accept <see cref="PocoType"/>'s descendant class instance
        ///     which is checked at compile time; however, whether exported collection item is actually of the descendant type
        ///     will only be checked at runtime during export (rather than configuration). If it cannot be cast to <typeparamref name="TA"/>,
        ///     null will be passed to the <paramref name="valueGetter"/> by <see cref="PocoColumnSource{TA,TV}"/>.
        /// </summary>
        /// <typeparam name="TA">
        ///     Instance type accepted by <paramref name="valueGetter"/>; must be same as <typeparamref name="TE"/> or its descendant.
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Source data type - type returned by <paramref name="valueGetter"/>.
        /// </typeparam>
        /// <param name="valueGetter">
        ///     Mandatory, reference returning property or field or arbitrary value. In the latter case
        ///     <paramref name="sheetColumnCaption"/> should be specified and it will be used as source column name
        ///     for identification. The expression is compiled, cached and used for retrieving values for the column.
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Column header text in excel.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        public PocoExportConfigurator<TE> AddColumnPolymorphic<TA, TV>(
            Expression<Func<TA, TV>> valueGetter
            , string sheetColumnCaption = null, bool? autoFit = null, string format = null)
            where   TA: class
        {
            Check.DoRequireArgumentNotNull(valueGetter, nameof(valueGetter));
            Check.DoCheckArgument(typeof(TA).IsAssignableFrom(PocoType) || PocoType.IsAssignableFrom(typeof(TA))
                , () => $"{PocoType.Name} cannot be cast to {typeof(TA).Name}");

            var columnSource = PocoColumnSourceFactory.Create(valueGetter, sheetColumnCaption);

            return AddColumn(columnSource, sheetColumnCaption, autoFit, format);
        }

        /// <summary>
        ///     Low level method allowing client to configure column source in any way they need.
        ///     Creates column next to the right most existing column.
        /// </summary>
        /// <param name="columnDataSource">
        ///     Encapsulates data retrieval for the column.
        /// </param>
        /// <param name="sheetColumnCaption">
        ///     Optional, column header text in excel.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        public PocoExportConfigurator<TE> AddColumn(IColumnDataSource columnDataSource, string sheetColumnCaption = null, bool? autoFit = null, string format = null)
        {
            AddColumn(columnDataSource, Config.Columns.Count, sheetColumnCaption, autoFit, format);

            return this;
        }

        /// <summary>
        ///     Configure export of collection member with up to <paramref name="columnCount"/> elements.
        ///     A column will be created for every collection item.
        /// </summary>
        /// <param name="memberName">
        ///     Name of property or field returning array or <see cref="IList{TV}"/>, mandatory.
        /// </param>
        /// <param name="columnCount">
        ///     Number of columns to create. Collection elements exceeding this limit will not be exported.
        /// </param>
        /// <param name="sheetColumnCaptionFormat">
        ///     Optional, default is same as <paramref name="memberName"/>; columns will be named by appending '[index]' to the base.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        public PocoExportConfigurator<TE> AddCollectionColumns<TV>(string memberName, int columnCount, string sheetColumnCaptionFormat, bool? autoFit = null, string format = null)
        {
            Check.DoRequireArgumentNotBlank(memberName, nameof(memberName));
            Check.DoCheckArgument(columnCount > 0, nameof(columnCount));

            sheetColumnCaptionFormat = sheetColumnCaptionFormat ?? memberName;

            var memberInfo = PocoType.GetMember(memberName).FirstOrDefault(i => i is PropertyInfo || i is FieldInfo);
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
                , () => $"{PocoType.Name}.{memberName} does not implement {typeof(IList<TV>).Name}.");

            AddCollectionColumns(collectionExtractor, columnCount, sheetColumnCaptionFormat, autoFit, format);

            return this;
        }

        /// <summary>
        ///     Configure export of generic list (or array) member with up to <paramref name="columnCount"/> elements.
        ///     A column will be created for every collection item, limited by max number supported by excel.
        /// </summary>
        /// <param name="collectionMemberGetter">
        ///     Name of property or field returning array or <see cref="IList{TV}"/>, mandatory. If 
        ///     Mandatory, reference returning property or field implementing <see cref="IList{TV}"/>.
        ///     If it is not a property or field reference, <paramref name="sheetColumnCaptionFormat"/> must be specified
        ///     and it will be used as source column name base for identification.
        ///     The expression is compiled, cached and used for retrieving values for the column.
        /// </param>
        /// <param name="columnCount">
        ///     Number of columns to create. Collection elements exceeding this limit will not be exported.
        ///     Max number of columns supported by excel is 16384.
        /// </param>
        /// <param name="sheetColumnCaptionFormat">
        ///     .Net format string accepting index as the only argument. Default is 'MemberName[{0}]' where MemberName
        ///     is <paramref name="collectionMemberGetter"/>'s name.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        public PocoExportConfigurator<TE> AddCollectionColumns<TV>(
            Expression<Func<TE, IList<TV>>> collectionMemberGetter, int columnCount, string sheetColumnCaptionFormat = null, bool? autoFit = null, string format = null)
        {
            return AddInheritedCollectionColumns(collectionMemberGetter, columnCount, sheetColumnCaptionFormat, autoFit, format);
        }

        /// <summary>
        ///     Configure export of generic list (or array) member with up to <paramref name="columnCount"/> elements.
        ///     A column will be created for every collection item, limited by max number supported by excel.
        ///     The expression <paramref name="collectionMemberGetter"/> can accept <see cref="PocoType"/>'s base class instance; in this
        ///     case type safety will be checked at runtime (during configuration) as it cannot be enforced at compile time.
        /// </summary>
        /// <typeparam name="TA">
        ///     Instance type accepted by <paramref name="collectionMemberGetter"/>; same as <typeparamref name="TE"/> or its base class
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Element type of the collection returned by <paramref name="collectionMemberGetter"/>.
        /// </typeparam>
        /// <param name="collectionMemberGetter">
        ///     Name of property or field returning array or <see cref="IList{TV}"/>, mandatory. If 
        ///     Mandatory, reference returning property or field implementing <see cref="IList{TV}"/>.
        ///     If it is not a property or field reference, <paramref name="sheetColumnCaptionFormat"/> must be specified
        ///     and it will be used as source column name base for identification.
        ///     The expression is compiled, cached and used for retrieving values for the column.
        /// </param>
        /// <param name="columnCount">
        ///     Number of columns to create. Collection elements exceeding this limit will not be exported.
        ///     Max number of columns supported by excel is 16384.
        /// </param>
        /// <param name="sheetColumnCaptionFormat">
        ///     .Net format string accepting index as the only argument. Default is 'MemberName[{0}]' where MemberName
        ///     is <paramref name="collectionMemberGetter"/>'s name.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        /// <remarks>
        ///     There should be really reverse type constraint 'TE: TA', TA should be base class for TE, but
        ///     I cannot find a way to express it.
        ///     This method encapsulate runtime (during configuration) type check.
        /// </remarks>
        public PocoExportConfigurator<TE> AddInheritedCollectionColumns<TA, TV>(
            Expression<Func<TA, IList<TV>>> collectionMemberGetter, int columnCount, string sheetColumnCaptionFormat = null, bool? autoFit = null, string format = null)
            where TA: class
        {
            Check.DoCheckArgument(typeof(TA).IsAssignableFrom(PocoType), () => $"{PocoType.Name} does not inherit from {typeof(TA).Name}");

            return AddCollectionColumnPolymorphic(collectionMemberGetter, columnCount, sheetColumnCaptionFormat, autoFit, format);
        }

        /// <summary>
        ///     Configure export of generic list (or array) member with up to <paramref name="columnCount"/> elements.
        ///     A column will be created for every collection item, limited by max number supported by excel.
        ///     The expression can accept <see cref="PocoType"/>'s descendant class instance which is checked at compile time;
        ///     however, whether exported collection item is actually of the descendant type
        ///     will only be checked at runtime during export (rather than configuration). If it cannot be cast to <typeparamref name="TA"/>,
        ///     null will be passed to the <paramref name="collectionMemberGetter"/> by <see cref="PocoColumnSource{TA,TV}"/>.
        /// </summary>
        /// <typeparam name="TA">
        ///     Instance type accepted by <paramref name="collectionMemberGetter"/>; same as <typeparamref name="TE"/> or its base class
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Element type of the collection returned by <paramref name="collectionMemberGetter"/>.
        /// </typeparam>
        /// <param name="collectionMemberGetter">
        ///     Name of property or field returning array or <see cref="IList{TV}"/>, mandatory. If 
        ///     Mandatory, reference returning property or field implementing <see cref="IList{TV}"/>.
        ///     If it is not a property or field reference, <paramref name="sheetColumnCaptionFormat"/> must be specified
        ///     and it will be used as source column name base for identification.
        ///     The expression is compiled, cached and used for retrieving values for the column.
        /// </param>
        /// <param name="columnCount">
        ///     Number of columns to create. Collection elements exceeding this limit will not be exported.
        ///     Max number of columns supported by excel is 16384.
        /// </param>
        /// <param name="sheetColumnCaptionFormat">
        ///     .Net format string accepting index as the only argument. Default is 'MemberName[{0}]' where MemberName
        ///     is <paramref name="collectionMemberGetter"/>'s name.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        /// <remarks>
        ///     This method encapsulates generic type constraint.
        /// </remarks>
        public PocoExportConfigurator<TE> AddDescendantsCollectionColumns<TA, TV>(
            Expression<Func<TA, IList<TV>>> collectionMemberGetter, int columnCount, string sheetColumnCaptionFormat = null, bool? autoFit = null, string format = null)
            where TA: class, TE
        {
            return AddCollectionColumnPolymorphic(collectionMemberGetter, columnCount, sheetColumnCaptionFormat, autoFit, format);
        }

        /// <summary>
        ///     Configure export of generic list (or array) member with up to <paramref name="columnCount"/> elements.
        ///     A column will be created for every collection item, limited by max number supported by excel.
        ///     The expression <paramref name="collectionMemberGetter"/> can accept <see cref="PocoType"/>'s base base or descendant
        ///     class instance; actual instances will be cast to <typeparamref name="TA"/> at run time and null will be passed
        ///     to <paramref name="collectionMemberGetter"/> if the cast fails.
        /// </summary>
        /// <typeparam name="TA">
        ///     Instance type accepted by <paramref name="collectionMemberGetter"/>; same as <typeparamref name="TE"/>, its base or descendant.
        /// </typeparam>
        /// <typeparam name="TV">
        ///     Element type of the collection returned by <paramref name="collectionMemberGetter"/>.
        /// </typeparam>
        /// <param name="collectionMemberGetter">
        ///     Name of property or field returning array or <see cref="IList{TV}"/>, mandatory. If 
        ///     Mandatory, reference returning property or field implementing <see cref="IList{TV}"/>.
        ///     If it is not a property or field reference, <paramref name="sheetColumnCaptionFormat"/> must be specified
        ///     and it will be used as source column name base for identification.
        ///     The expression is compiled, cached and used for retrieving values for the column.
        /// </param>
        /// <param name="columnCount">
        ///     Number of columns to create. Collection elements exceeding this limit will not be exported.
        ///     Max number of columns supported by excel is 16384.
        /// </param>
        /// <param name="sheetColumnCaptionFormat">
        ///     .Net format string accepting index as the only argument. Default is 'MemberName[{0}]' where MemberName
        ///     is <paramref name="collectionMemberGetter"/>'s name.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        /// <remarks>
        ///     There should be really reverse type constraint 'TE: TA', TA should be base class for TE, but
        ///     I cannot find a way to express it.
        /// </remarks>
        public PocoExportConfigurator<TE> AddCollectionColumnPolymorphic<TA, TV>(
            Expression<Func<TA, IList<TV>>> collectionMemberGetter
            , int columnCount
            , string sheetColumnCaptionFormat = null, bool? autoFit = null, string format = null)
            where TA: class
        {
            Check.DoRequireArgumentNotNull(collectionMemberGetter, nameof(collectionMemberGetter));
            Check.DoCheckArgument(typeof(TA).IsAssignableFrom(PocoType) || PocoType.IsAssignableFrom(typeof(TA))
                , () => $"{PocoType.Name} cannot be cast to {typeof(TA).Name}");

            Check.DoCheckArgument(columnCount > 0 && columnCount < 16384, nameof(columnCount));

            var memberInfo = ExpressionHelper.GetMember(collectionMemberGetter);

            Check.DoCheckArgument(memberInfo != null || !string.IsNullOrWhiteSpace(sheetColumnCaptionFormat)
                                  , "If collection expression does not refer to property or field, column caption format must be provided");

            if (string.IsNullOrWhiteSpace(sheetColumnCaptionFormat))
            {
                Debug.Assert(memberInfo != null, nameof(memberInfo) + " != null");
                sheetColumnCaptionFormat = $"{memberInfo.Name}[{{0}}]";
            }

            var memberName = memberInfo?.Name ?? collectionMemberGetter.ToString();

            var propertyInfo = memberInfo as PropertyInfo;
            Type collectionType;
            var collectionExtractor = LambdaExpressionCache.Compile(collectionMemberGetter);
            if (propertyInfo != null)
            {
                Check.DoCheckArgument(propertyInfo.CanRead, () => $"Property {memberName} of {PocoType.Name} is not readable");
                collectionType = propertyInfo.PropertyType;
            }
            else
            {
                var fieldInfo = (FieldInfo)memberInfo;
                Debug.Assert(fieldInfo != null, nameof(fieldInfo) + " != null");
                collectionType = fieldInfo.FieldType;
            }

            Check.DoCheckArgument(typeof(IList<TV>).IsAssignableFrom(collectionType)
                , () => $"{PocoType.Name}.{memberName} is not a generic list.");

            AddCollectionColumns(collectionExtractor, columnCount, sheetColumnCaptionFormat, autoFit, format);

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
        ///     Column header text in excel. Optional if <paramref name="columnDataSource"/> has non-null virtual column name.
        ///     Empty string allowed.
        /// </param>
        /// <param name="autoFit">
        ///     Optional, whether to fit column width to content after export (default is true).
        /// </param>
        /// <param name="format">
        ///     Optional format for excel cells.
        /// </param>
        /// <returns>
        ///     Itself
        /// </returns>
        public DataColumnExportAutoConfig AddColumn(IColumnDataSource columnDataSource, int sheetColumnIndex, string sheetColumnCaption, bool? autoFit, string format)
        {
            Check.DoRequireArgumentNotNull(columnDataSource, nameof(columnDataSource));
            Check.DoCheckArgument(sheetColumnCaption != null || columnDataSource.Name != null, "Non-null column caption cannot be resolved.");
            Check.DoCheckArgument(sheetColumnIndex >= 0, "Column index cannot be negative.");

            if (string.IsNullOrEmpty(sheetColumnCaption) && columnDataSource.Name != null)
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

        private void AddCollectionColumns<TA, TV>(Func<TA, IList<TV>> collectionGetter,
            int columnCount,
            string sheetColumnCaptionFormat,
            bool? autoFit = null,
            string format = null)
            where TA: class
        {
            Check.DoCheckArgument(typeof(TA).IsAssignableFrom(PocoType), () => $"{PocoType.Name} does not inherit from {typeof(TA).Name}");

            for (var i = 0; i < columnCount; ++i)
            {
                var collectionIndex = i;
                var columnName = string.Format(sheetColumnCaptionFormat, collectionIndex);
                var columnSource = new PocoColumnSource<TA, TV>(columnName, o => TryGetCollectionElement(collectionGetter(o), collectionIndex));
                AddColumn(columnSource, Config.Columns.Count, columnName, autoFit, format);
            }
        }
    }
}