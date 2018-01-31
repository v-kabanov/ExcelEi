using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelEi.Read;
using NUnit.Framework;
using OfficeOpenXml;

namespace ExcelEi.Test
{
    public class NameValue
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }

    public class NameValueMap
    {
        public NameValueMap()
        {
        }
    }
    /// <summary>
    ///     Strongly typed entry in common wafer defect data spreadsheet, represents 1 'valid' die defect occurrence.
    /// </summary>
    public class DefectRecord
    {
        public string Valid { get; set; }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        /// <summary>
        ///     Optional, not provided in Die Crack MOI
        /// </summary>
        public double DefectAreaSize { get; set; }
        public int DefectCode { get; set; }
    }

    public class DefectReader : TableMappingReader<DefectRecord>
    {
        public DefectReader(bool readArea)
        {
            Map(r => r.Valid);
            Map(r => r.RowIndex, "Row");
            Map(r => r.ColumnIndex, "Col");
            Map(r => r.DefectCode, "ClassCd", ConvertDefectCode);

            if (readArea)
                Map(r => r.DefectAreaSize, "Area");
        }

        private int ConvertDefectCode(object value)
        {
            var stringValue = value as string;
            if (!string.IsNullOrWhiteSpace(stringValue) && stringValue.Length == 1 && !char.IsDigit(stringValue[0]))
            {
                return stringValue[0];
            }
            return Conversion.GetTypedExcelValue<int>(value);
        }
    }

    [TestFixture]
    public class ExcelReadingTest
    {
        private const string TestFileWithTablesRelativePath = @"TestData\testbook1.tables.xlsx";
        private const string TestFileRelativePath = @"TestData\testbook1.xlsx";
        private const string XlsTestFileRelativePath = @"TestData\testbook1.xls";
        private const string ExcelTableNameSummary = "tbSummary";
        private const string ExcelTableNameData = "tbData";

        public string TestFilePath { get; }

        public string TestFileWithTablesPath { get; }

        public string XlsTestFilePath { get; }

        public ExcelReadingTest()
        {
            var testDirectory = TestContext.CurrentContext.TestDirectory;

            TestFilePath = Path.Combine(testDirectory, TestFileRelativePath);
            TestFileWithTablesPath = Path.Combine(testDirectory, TestFileWithTablesRelativePath);
            XlsTestFilePath = Path.Combine(testDirectory, XlsTestFileRelativePath);
        }


        [Test]
        public void ExcelTableAccessibleAsNamedRange()
        {
            using (var package = new ExcelPackage(new FileInfo(TestFileWithTablesPath)))
            {
                var workbook = package.Workbook;

                var tbSummary = workbook.Worksheets.SelectMany(sh => sh.Tables).FirstOrDefault(n => n.Name == ExcelTableNameSummary);
                var tbData = workbook.Worksheets.SelectMany(sh => sh.Tables).FirstOrDefault(n => n.Name == ExcelTableNameData);

                var startRowIndex = tbData.Address.Start.Row;
                if (tbData.ShowFilter)
                    ++startRowIndex;

                for (var rowIndex = startRowIndex; rowIndex <= tbData.Address.End.Row; ++ rowIndex)
                {
                    Console.WriteLine("1: {0}", tbData.WorkSheet.Cells[rowIndex, 1].Value);
                }
            }
        }

        [Test]
        public void ExcelTableReaderTest()
        {
            using (var package = new ExcelPackage(new FileInfo(TestFileWithTablesPath)))
            {
                var workbook = package.Workbook;

                var tbSummary = workbook.Worksheets.SelectMany(sh => sh.Tables).FirstOrDefault(n => n.Name == ExcelTableNameSummary);
                var tbData = workbook.Worksheets.SelectMany(sh => sh.Tables).FirstOrDefault(n => n.Name == ExcelTableNameData);

                var reader = new ExcelTableReader(tbData);

                Trace(new ExcelTableReader(tbSummary));
                Trace(reader);
            }
        }

        [Test]
        public void ReadAdHocTableXlsx()
        {
            IList<IDictionary<string, object>> epplusResult;

            using (var package = new ExcelPackage(new FileInfo(TestFilePath)))
            {
                var workbook = package.Workbook;
                var epplusReader = ExcelTableReader.ReadContiguousTableWithHeader(workbook.Worksheets[1], 17);
                //Trace(reader);
                epplusResult = epplusReader.Rows.ToListOfDictionaries();
            }

            var reader = AdoTableReader.ReadContiguousExcelTableWithHeader(TestFilePath, null, 17);
            //Trace(reader);
            var excelDataReaderResult = reader.Rows.ToListOfDictionaries();


            Assert.AreEqual(1895, excelDataReaderResult.Count);
            Assert.AreEqual(epplusResult.Count, excelDataReaderResult.Count);
            Assert.AreEqual(epplusResult, excelDataReaderResult);
        }

        [Test]
        public void ReadAdHocTableExcelDataReader()
        {
            ITableReader reader = AdoTableReader.ReadContiguousExcelTableWithHeader(XlsTestFilePath, null, 17);
            //Trace(reader);

            var xlsResult = reader.Rows.ToListOfDictionaries();

            reader = AdoTableReader.ReadContiguousExcelTableWithHeader(TestFilePath, null, 17);
            //Trace(reader);

            var xlsxResult = reader.Rows.ToListOfDictionaries();

            Assert.AreEqual(xlsResult, xlsxResult);

            Assert.AreEqual(1895, xlsResult.Count);
        }

        [Test]
        public void ReadArbitraryTable()
        {
            IList<IDictionary<string, object>> epplusResult;
            IList<IDictionary<string, object>> excelDataReaderResult;

            var columns = new List<KeyValuePair<string, int>>
                {
                    new KeyValuePair<string, int>("Name", 1),
                    new KeyValuePair<string, int>("Value", 9),
                };

            using (var package = new ExcelPackage(new FileInfo(TestFilePath)))
            {
                var workbook = package.Workbook;
                var reader = ExcelTableReader.ReadArbitraryTable(workbook.Worksheets[1], 7, 16, columns);
                //Trace(reader);
                TestReadMappedSummaryTable(reader);
                epplusResult = reader.Rows.ToListOfDictionaries();
            }

            {
                var reader = AdoTableReader.ReadArbitraryExcelTable(TestFilePath, null, 7, 16, columns);

                TestReadMappedSummaryTable(reader);
                excelDataReaderResult = reader.Rows.ToListOfDictionaries();
            }

            Assert.AreEqual(epplusResult, excelDataReaderResult);
        }


        [Test]
        public void AdHocTest()
        {
            using (var package = new ExcelPackage(new FileInfo(TestFilePath)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[1];

                var dataTableReader = ExcelTableReader.ReadContiguousTableWithHeader(worksheet, 15);

                var defectReader = new DefectReader(false);

                var result = defectReader.Read(dataTableReader);
                Assert.IsNotEmpty(result);
            }
        }

        private static void TestReadMappedSummaryTable(ITableReader tableReader)
        {
            var dict = tableReader.Rows
                .Where(r => r["Name"] != null)
                .ToDictionary(r => r.GetValue<string>("Name"), r => r["Value"]);
            Assert.AreEqual(7, dict.Count);

            var normalReader = new TableMappingReader<NameValue>()
                .Map(o => o.Name)
                .Map(o => o.Value);

            var nameValueList = normalReader.Read(tableReader);
            Assert.GreaterOrEqual(nameValueList.Count, 2);
            Assert.IsTrue(nameValueList.Any(o => o.Name != o.Value));

            var nameNameReader = new TableMappingReader<NameValue>()
                .Map(o => o.Name)
                .Map(o => o.Value, "Name");

            nameValueList = nameNameReader.Read(tableReader);
            Assert.GreaterOrEqual(nameValueList.Count, 2);
            Assert.GreaterOrEqual(nameValueList.Count(o => !string.IsNullOrEmpty(o.Name)), 2);
            Assert.IsFalse(nameValueList.Any(o => o.Name != o.Value));

            nameValueList = normalReader.Read(tableReader);
            Assert.GreaterOrEqual(nameValueList.Count, 2);
            Assert.IsTrue(nameValueList.Any(o => o.Name != o.Value));
        }

        private void Trace(ITableReader reader)
        {
            foreach (var row in reader.Rows)
            {
                foreach (var columnName in reader.Columns)
                {
                    Console.Write(row[columnName]);
                    Console.Write("\t");
                }
                Console.WriteLine();
            }
        }
    }
}