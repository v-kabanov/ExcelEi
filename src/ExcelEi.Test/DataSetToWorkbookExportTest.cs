// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-19
// Comment		
// **********************************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using NUnit.Framework;
using ExcelEi.Read;
using OfficeOpenXml;
using System.Linq;
using System.Linq.Expressions;
using System.Threading;
using ExcelEi.Write;
using NUnit.Framework.Internal;

namespace ExcelEi.Test
{
    public class PocoOne
    {
        private static int _idSequence;
        private static readonly Random Random = new Random();

        public PocoOne(int valueCount)
        {
            Values = new double?[valueCount];

            for (var i = 0; i < valueCount; ++i)
            {
                if (Random.NextDouble() < 0.1)
                    Values[i] = null;
                else
                    Values[i] = Random.NextDouble() * 500;
            }
        }

        public int Id { get; } = Interlocked.Increment(ref _idSequence);

        public DateTime DateTime = DateTime.Now;

        public double?[] Values { get; set; }
    }

    [TestFixture]
    public class DataSetToWorkbookExportTest
    {
        private static string TestFileDirectoryPath => TestContext.CurrentContext.TestDirectory;

        [Test]
        public void OneTable()
        {
            var file = new FileInfo(Path.Combine(TestFileDirectoryPath, Path.ChangeExtension(Guid.NewGuid().ToString(), "xlsx")));
            var workbook = new ExcelPackage(file);

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
        }

        [Test]
        public void ArrayFromPoco()
        {
            var file = new FileInfo(Path.Combine(TestFileDirectoryPath, Path.ChangeExtension(Guid.NewGuid().ToString(), "xlsx")));
            var workbook = new ExcelPackage(file);

            var dataSetExportConfig = new DataSetExportAutoConfig();

            var configurator = new PocoExportConfigurator(typeof(PocoOne), "OneSheet", "One");

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

            configurator = new PocoExportConfigurator(typeof(PocoOne), "TwoSheet");

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
            workbook.Dispose();
        }
    }
}
