// /**********************************************************************************************
// Author:  Vasily Kabanov
// Created  2018-01-09
// Comment  
// **********************************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelEi.Read
{
    public class DataSetExtractor
    {
        public IDictionary<string, ITableDataSource> TableDataSourceDescriptors { get; }

        public DataSetExtractor()
        {
            TableDataSourceDescriptors = new Dictionary<string, ITableDataSource>(StringComparer.OrdinalIgnoreCase);
        }

        public void AddTable(ITableDataSource dataSourceDescriptor, string tableName)
        {
            Check.DoRequireArgumentNotNull(dataSourceDescriptor, nameof(dataSourceDescriptor));
            Check.DoRequireArgumentNotBlank(tableName, nameof(tableName));

            TableDataSourceDescriptors.Add(tableName, dataSourceDescriptor);
        }

        public DataSet ExtractDataSet(IDictionary<string, IEnumerable> dataSources)
        {
            var result = new DataSet();

            foreach (var dataSourcePair in dataSources)
            {
                var dataSourceDescriptor = TableDataSourceDescriptors.TryGetValue(dataSourcePair.Key);
                Check.DoEnsureLambda(dataSourceDescriptor != null, () => $"No descriptor for {dataSourcePair.Key}");
                var table = ExtractTable(dataSourceDescriptor, dataSourcePair.Value);
                table.TableName = dataSourcePair.Key;

                result.Tables.Add(table);
            }

            return result;
        }

        public static DataTable ExtractTable(ITableDataSource dataSource, IEnumerable data)
        {
            var result = new DataTable();
            var columnDescriptors = dataSource.AllColumns.ToArray();
            foreach (var sourceColumn in columnDescriptors)
            {
                result.Columns.Add(sourceColumn.Name, sourceColumn.DataType);
            }

            foreach (var dataItem in data)
            {
                var row = result.NewRow();

                for (var columnIndex = 0; columnIndex < columnDescriptors.Length; ++columnIndex)
                {
                    row[columnIndex] = columnDescriptors[columnIndex].GetValue(dataItem) ?? DBNull.Value;
                }
            }

            return result;
        }
    }
}