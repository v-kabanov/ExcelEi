// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-19
// Comment		
// **********************************************************************************************/

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using NUnit.Framework;
using ExcelEi.Read;
using OfficeOpenXml;
using System.Linq;
using System.Linq.Expressions;
using System.Threading;
using ExcelEi.Write;

namespace ExcelEi.Test
{
    public class PocoOne
    {
        private static int _idSequence;
        private static readonly Random Random = new Random();

        public PocoOne()
        {
        }

        public PocoOne(int valueCount)
        {
            Id = Interlocked.Increment(ref _idSequence);

            DateTime = DateTime.Now;

            Values = new double?[valueCount];

            for (var i = 0; i < valueCount; ++i)
            {
                if (Random.NextDouble() < 0.1)
                    Values[i] = null;
                else
                    Values[i] = Random.NextDouble() * 500;
            }
        }

        public int Id { get; set; }

        public DateTime DateTime { get; set; }

        public double?[] Values { get; set; }
    }

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

    [TestFixture]
    public class DataSetToWorkbookExportTest
    {
        private string GetNewOutFilePath() =>
            Path.Combine(TestContext.CurrentContext.WorkDirectory, $"excelei-{DateTime.Now:MM-dd-HHmmss}.xlsx");

        [Test]
        public void OneTable()
        {
            var outPath = GetNewOutFilePath();
            var workbook = new ExcelPackage(new FileInfo(outPath));

            var dataSet = new DataSet();
            var tableName = "LastSegments";
            var dataTable = dataSet.Tables.Add(tableName);
            dataTable.Columns.Add("Id", typeof(long));
            dataTable.Columns.Add("Description", typeof(string));
            dataTable.Columns.Add("Authority", typeof(string));
            dataTable.Columns.Add("MaterialLotID", typeof(string));
            dataTable.Columns.Add("LotType", typeof(string));
            dataTable.Columns.Add("Designation", typeof(string));
            dataTable.Columns.Add("SegmentState", typeof(string));

            var designations = new [] {"Production", "R&D", "Baseline"};
            var segmentStates = new[] {"Commenced", "Completed", "Aborted"};
            var rowCount = 1000;

            var random = new Random();

            for (var n = 0; n < rowCount; ++n)
            {
                dataTable.Rows.Add(
                    rowCount * 5 - n
                    , $"Description{n}"
                    , $"Authority{n % 3}"
                    , $"Lot#{n}"
                    , $"LotType{n%2}"
                    , designations[random.Next(0, designations.Length)]
                    , segmentStates[random.Next(0, segmentStates.Length)]);
            }

            var config = new DataSetExportAutoConfig(dataSet);

            var tableConfig = config.GetTableConfig(tableName);
            var columnConfig = tableConfig.GetAutoColumnConfig("Description");
            columnConfig.MinimumWidth = 40;
            columnConfig.WrapText = true;

            tableConfig.SheetName = "Last Segments";

            tableConfig.GetAutoColumnConfig("Authority").AutoFit = true;
            tableConfig.GetAutoColumnConfig("MaterialLotID").AutoFit = true;
            tableConfig.GetAutoColumnConfig("LotType").AutoFit = true;
            tableConfig.GetAutoColumnConfig("Designation").AutoFit = true;

            columnConfig = tableConfig.GetAutoColumnConfig("SegmentState");
            columnConfig.Caption = "Segment State";
            columnConfig.AutoFit = true;
            columnConfig.BackgroundColorExtractor = (d, n) => "Aborted".Equals(((DataRow)d)["SegmentState"]) ? Color.LightCoral : (Color?)null;

            var exporter = new DataSetToWorkbookExporter(config) {DataSet = new DataSetAdapter(dataSet)};

            var start = DateTime.Now;

            exporter.Export(workbook);

            var duration = DateTime.Now - start;

            Console.WriteLine("Duration: {0}", duration);

            workbook.Save();

            TestContext.WriteLine($"Saved {outPath}.");

            workbook.Dispose();
            var readTable = AdoTableReader.GetWorksheetDataTable(outPath);

            Assert.AreEqual(tableConfig.Columns.Count, readTable.Columns.Count);

            for (var i = 0; i < tableConfig.Columns.Count; ++i)
            {
                //ExcelDataReader does not set column name, caption and contains header as first data row
                Assert.AreEqual(tableConfig.Columns[i].Caption, readTable.Rows[0][i]);
                // since header is first data row, only string type will be set correctly
                //Assert.AreEqual(tableConfig.Columns[i].ColumnDataSource.DataType, readTable.Columns[i].DataType);
            }

            Assert.AreEqual(rowCount, readTable.Rows.Count - 1);

            for (int i = 0; i < rowCount; ++i)
            {
                var savedRow = dataTable.Rows[i];
                var readRow = readTable.Rows[i + 1];
                for (var c = 0; c < tableConfig.Columns.Count; ++c)
                {
                    Assert.AreEqual(savedRow[c], readRow[c]);
                }
            }

            File.Delete(outPath);
        }

        [Test]
        public void ArrayFromPoco()
        {
            var outPath = GetNewOutFilePath();
            var workbook = new ExcelPackage(new FileInfo(outPath));

            var dataSetExportConfig = new DataSetExportAutoConfig();

            var configurator = new PocoExportConfigurator<PocoOne>("OneSheet", "One");

            Expression<Func<PocoOne, int>> refId = o => o.Id;
            Expression<Func<PocoOne, DateTime>> refDateTime = o => o.DateTime;
            Expression<Func<PocoOne, IList<double?>>> refCollection = o => o.Values;
            Expression<Func<PocoOne, string>> refJoinedCollection = o => o.Values != null ? string.Join(",", o.Values.Select(e => e.ToString())) : null;

            configurator
                .AddColumn(refId)
                .AddColumn(refDateTime)
                .AddColumn(refJoinedCollection, "Joined Values");

            configurator.AddCollectionColumns(refCollection, 5);

            dataSetExportConfig.AddSheet(configurator.Config);

            configurator = new PocoExportConfigurator<PocoOne>("TwoSheet");

            configurator.AddColumn(refDateTime);

            configurator.AddCollectionColumns(refCollection, 10);

            configurator.AddColumn(refId);

            dataSetExportConfig.AddSheet(configurator.Config);

            var dataSet = new DataSetAdapter();
            var data1 = Enumerable.Range(0, 100)
                .Select(i => new PocoOne(6))
                .ToList();
            var data2 = Enumerable.Range(0, 1000)
                .Select(i => new PocoOne(9))
                .ToList();

            dataSet.Add(data1, "One");
            dataSet.Add(data2, "TwoSheet");

            var exporter = new DataSetToWorkbookExporter(dataSetExportConfig) {DataSet = dataSet};

            exporter.Export(workbook);

            workbook.Save();
            TestContext.WriteLine($"Saved {outPath}.");

            workbook.Dispose();

            workbook = new ExcelPackage(new FileInfo(outPath));

            var reader = ExcelTableReader.ReadContiguousTableWithHeader(workbook.Workbook.Worksheets[1], 1);
            var readPocos = new PocoOneReader().Read(reader);

            Assert.AreEqual(data1.Count, readPocos.Count);

            for (int i = 0; i < data1.Count; ++i)
            {
                var saved = data1[i];
                var read = readPocos[i];
                Assert.AreEqual(saved.Id, read.Id);
                Assert.Less((saved.DateTime - read.DateTime).TotalMilliseconds, 1);
                Assert.AreEqual(saved.Values.Length, read.Values.Length);

                for (var j = 0; j < saved.Values.Length; ++j)
                {
                    if (saved.Values[j].HasValue)
                        Assert.AreEqual(saved.Values[j].Value, read.Values[j], 0.0000001D);
                    else
                        Assert.IsFalse(read.Values[j].HasValue);
                }
            }

            workbook.Dispose();

            File.Delete(outPath);
        }
    }
}
