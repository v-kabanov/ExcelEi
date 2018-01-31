using System.Collections.Generic;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Component exposing row readers accessible via 0-based index.
    /// </summary>
    public interface ITableRowReaderCollection : IEnumerable<ITableRowReader>
    {
        /// <summary>
        ///     Null when created as scanning reader estimating where table ends by all cells having null value.
        /// </summary>
        int? Count { get; }

        /// <summary>
        ///     Get reader for row by its index.
        /// </summary>
        /// <param name="rowIndex">
        ///     Relative zero-bound row index
        /// </param>
        ITableRowReader this[int rowIndex] { get; }
    }
}