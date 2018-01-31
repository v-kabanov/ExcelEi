// /**********************************************************************************************
// Author:		Vasily Kabanov
// Created		2015-02-18
// Comment		
// **********************************************************************************************/

using ExcelEi.Read;
using OfficeOpenXml;

namespace ExcelEi.Write
{
    /// <summary>
    ///     Implements export from <see cref="IDataSet"/> into <see cref="ExcelPackage"/> (Excel workbook).
    /// </summary>
    public class DataSetToWorkbookExporter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="T:System.Object"/> class.
        /// </summary>
        public DataSetToWorkbookExporter(IWorkbookExportConfig exportConfig)
        {
            Check.DoRequireArgumentNotNull(exportConfig, nameof(exportConfig));

            ExportConfig = exportConfig;
        }

        /// <summary>
        ///     Get object defining how to export data.
        /// </summary>
        public IWorkbookExportConfig ExportConfig { get; }

        /// <summary>
        ///     Get or set data to be exported.
        /// </summary>
        public IDataSet DataSet { get; set; }

        public void Export(ExcelPackage excelPackage)
        {
            Check.DoRequireArgumentNotNull(excelPackage, "excelPackage");

            foreach (var sheetConfig in ExportConfig.SheetTables)
            {
                var sheet = excelPackage.Workbook.Worksheets.Add(sheetConfig.SheetName);
                var dataTable = DataSet.DataTables[sheetConfig.DataTableName];

                var exporter = new DataTableToSheetExporter(sheet, sheetConfig, dataTable);
                exporter.Export();
            }
        }
    }
}
