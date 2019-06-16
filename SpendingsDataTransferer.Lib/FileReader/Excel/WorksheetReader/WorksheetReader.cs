using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using SpendingsDataTransferer.Lib.ApplicationModel.Excel;

namespace SpendingsDataTransferer.Lib.FileReader.Excel.WorksheetReader
{
    public class WorksheetReader : IWorksheetReader
    {
        private readonly string filename;
        public WorksheetReader(string filename)
        {
            this.filename = filename;
        }

        public IEnumerable<Worksheet> Worksheets
        {
            get
            {
                var existingSpreadsheet = new FileInfo(filename);
                using(ExcelPackage package = new ExcelPackage(existingSpreadsheet))
                {
                    var result = package.Workbook.Worksheets
                        .Select(worksheet => new Worksheet(worksheet.Name)).ToArray();
                    return result;

                }
            }
        }
    }
}