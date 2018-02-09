# ExcelEi
Library for excel export-import

It builds on top of EPPlus, ExcelDataReader and AutoMapper.

Import example:
```C#
    public class NameValue
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }

    public void Foo()
    {
        var excelFilePath = "c:\\file.xlsx";
        var package = new ExcelPackage(new FileInfo(excelFilePath));

        var tableReader = ExcelTableReader.ReadContiguousTableWithHeader(package.Workbook.Worksheets[1], 1);

        var pocoReader = new TableMappingReader<NameValue>()
            .Map(o => o.Name)
            .Map(o => o.Value);

        IList<NameValue> pocoList = pocoReader.Read(tableReader);
    }

```
Export example:
```C#

        var outPath = "c:\\out.xlsx";
        var package = new ExcelPackage(new FileInfo(outPath));

        var dataSetExportConfig = new DataSetExportAutoConfig();
        
        // data mappings
        var configurator = new PocoExportConfigurator<PocoOne>("OneSheet", "One");

        Expression<Func<PocoOne, int>> refId = o => o.Id;
        Expression<Func<PocoOne, DateTime>> refDateTime = o => o.DateTime;
        Expression<Func<PocoOne, IList<double?>>> refCollection = o => o.Values;
        // calculated column
        Expression<Func<PocoOne, string>> refJoinedCollection = o => o.Values != null ? string.Join(",", o.Values.Select(e => e.ToString())) : null;

        configurator
            .AddColumn(refId)
            .AddColumn(refDateTime)
            .AddColumn(refJoinedCollection, "Joined Values");

        // will produce 5 columns for 5 of the first elements of the list, named 'Values[<index>]'
        configurator.AddCollectionColumns(refCollection, 5);

        dataSetExportConfig.AddSheet(configurator.Config);

        // data to be exported
        var dataSet = new DataSetAdapter();
        var data1 = Enumerable.Range(0, 100)
            .Select(i => new PocoOne(6))
            .ToList();

        dataSet.Add(data1, "One");

        // exporter has data mappings and formatting settings
        var exporter = new DataSetToWorkbookExporter(dataSetExportConfig) {DataSet = dataSet};

        exporter.Export(package);

        package.Save();
        // exporter and data mappings can be reused for other data sets
```

POCO readers can convert column values:
```C#
    public class PocoOneReader : TableMappingReader<PocoOne>
    {
        /// <inheritdoc />
        public PocoOneReader()
        {
            Map(o => o.Id);
            Map(o => o.DateTime);
            Map(o => o.Values, "Joined Values", ParseJoinedValues);
        }

        private double?[] ParseJoinedValues(object joinedValue)
        {
            return joinedValue
                ?.ToString()
                .Split(',')
                .Select(Parse)
                .ToArray();
        }

        private static double? Parse(string value)
        {
            if (string.IsNullOrEmpty(value))
                return null;

            return double.Parse(value);
        }
    }

```

Easy export from ADO.NET:
```C#

    var dataSet = new DataSet();
    // ...
    var config = new DataSetExportAutoConfig(dataSet);

    var exporter = new DataSetToWorkbookExporter(config) {DataSet = new DataSetAdapter(dataSet)};

    using(var package = new ExcelPackage(new FileInfo("c:\\output.xlsx")))
    {
        exporter.Export(package);
        package.Save();
    }
```

Supports reflection for difficult situations:
```C#
    configurator.AddColumn<int?>(nameof(PocoTwo.FieldInt), "Reflected Field");
```

Mappings for base classes can be encapsulated and reused with generics and inheritance (see wiki for more details).
Export:
```C#
    public class PocoBaseExportConfigurator<T> : PocoExportConfigurator<T>
        where T: PocoBase
    {
        public PocoBaseExportConfigurator(string sheetName, string dataTableName)
            : base(sheetName, dataTableName)
        {
            AddColumn(o => o.Id);
            AddColumn(o => o.DateTime);
        }
    }

    public class PocoOneExportBaseConfigurator<T> : PocoBaseExportConfigurator<T>
        where T: PocoOne
    {
        public PocoOneExportBaseConfigurator(string sheetName, string dataTableName)
            : base(sheetName, dataTableName)
        {
            AddColumn(o => o.Values != null ? string.Join(",", o.Values.Select(e => e.ToString())) : null, "Joined Values");
            AddCollectionColumns(o => o.Values, 5, "value#{0}");
        }
    }
```
Import:
```C#
    public class PocoOneBaseReader<T> : TableMappingReader<T>
        where T: PocoOne
    {
        public PocoOneBaseReader()
        {
            Map(o => o.Id);
            Map(o => o.DateTime);
        }
    }

    public class PocoTwoReader : PocoOneBaseReader<PocoTwo>
    {
        public PocoTwoReader()
        {
            Map(o => o.IntegerFromPocoTwo);
        }
    }
```

Supports sparse columns and custom placement (see also corresponding test):
```C#
    var exportConfig = new PocoThreeExportConfigurator(sheetName).Config;

    const int firstColumnIndex = 2;
    // this is index of the header row
    const int firstRowIndex = 3;

    exportConfig.LeftSheetColumnIndex = firstColumnIndex;
    exportConfig.TopSheetRowIndex = firstRowIndex;
    // no freezing panes
    exportConfig.FreezeColumnIndex = null;

    // move third column to the right
    Assert.IsNotEmpty(exportConfig.Columns[2].Caption, "Sheet column#2 has no caption");
    var movedColumnConfig = exportConfig.GetAutoColumnConfig(exportConfig.Columns[2].Caption);
    Assert.IsNotNull(movedColumnConfig, "Failed to find column export config by caption");
    movedColumnConfig.Index = exportConfig.Columns.Count + 2;
    // allow it to grow more at the end of the table
    movedColumnConfig.MaximumWidth = 300;

    // ...
    var columnReadingMap = exportConfig.Columns
        .Select(c => new KeyValuePair<string, int>(c.Caption, firstColumnIndex + c.Index))
        .ToList();

    const int startDataRowIndex = firstRowIndex + 1;
    var reader = new ExcelTableReader(workbook.Workbook.Worksheets[1], startDataRowIndex, null, columnReadingMap);
    var readPocos = new PocoThreeReader().Read(reader);

```
