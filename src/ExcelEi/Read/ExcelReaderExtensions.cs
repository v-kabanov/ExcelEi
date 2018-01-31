// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-04-05
// Comment		
// **********************************************************************************************/

using System.Collections.Generic;
using System.Linq;

namespace ExcelEi.Read
{
    public static class ExcelReaderExtensions
    {
        public static IDictionary<string, object> ToDictionary(this ITableRowReader rowReader)
        {
            Check.DoRequireArgumentNotNull(rowReader, nameof(rowReader));

            return rowReader.AllColumnNames.ToDictionary(n => n, n => rowReader[n]);
        }

        public static IList<IDictionary<string, object>> ToListOfDictionaries(this ITableRowReaderCollection rowReaderCollection)
        {
            Check.DoRequireArgumentNotNull(rowReaderCollection, nameof(rowReaderCollection));

            return rowReaderCollection.Select(r => r.ToDictionary()).ToList();
        }
    }
}