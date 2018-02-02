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
