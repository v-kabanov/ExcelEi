// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2017-03-20
// Comment		
// **********************************************************************************************/

using System.Collections;

namespace ExcelEi.Read
{
    public class PocoTableAdapter : IDataTable
    {
        IEnumerable Collection { get; }

        public PocoTableAdapter(IEnumerable collection)
        {
            Check.DoRequireArgumentNotNull(collection, nameof(collection));

            Collection = collection;
        }

        public IEnumerable Rows => Collection;
    }
}