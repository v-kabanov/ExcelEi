using System.Collections.Generic;

namespace ExcelEi.Read
{
    /// <summary>
    ///     Minimal functionality to be implemented by a component reading a table.
    /// </summary>
    public interface ITableReader
    {
        ITableRowReaderCollection Rows { get; }

        ICollection<string> Columns { get; }
    }
}