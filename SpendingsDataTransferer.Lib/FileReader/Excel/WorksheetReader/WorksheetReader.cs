using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using SpendingsDataTransferer.Lib.ApplicationModel.Excel;

namespace SpendingsDataTransferer.Lib.FileReader.Excel.WorksheetReader
{
    public class WorksheetReader : IWorksheetReader
    {
        private const string dateTimeFormat = "MM-dd-yyyy HH:mm";
        private readonly FileInfo existingSpreadsheet;
        public WorksheetReader(string filename)
        {
            existingSpreadsheet = new FileInfo(filename);
        }

        public IEnumerable<Worksheet> Worksheets
        {
            get
            {
                using(ExcelPackage package = new ExcelPackage(existingSpreadsheet))
                {
                    return package.Workbook.Worksheets
                        .Select(worksheet => new Worksheet(worksheet.Name)).ToArray();
                }
            }
        }

        public T GetCellValue<T>(int worksheetIndex, int rowIndex, int columnIndex)
        {
            using(ExcelPackage package = new ExcelPackage(existingSpreadsheet))
            {
                ThrowErrorIfRowIsLesserThanZero(rowIndex);
                ThrowErrorIfColumnIsLesserThanZero(columnIndex);

                var worksheets = package.Workbook.Worksheets;
                ThrowErrorIfTriedToReadNonExistentWorksheetIsLesserThanZero(worksheetIndex, worksheets.Count);

                if (typeof(T) == typeof(string))
                {
                    var cell = worksheets[worksheetIndex].Cells[rowIndex, columnIndex];
                    return GetStringValueOfCell<T>(cell);
                }
                
                return worksheets[worksheetIndex].GetValue<T>(rowIndex, columnIndex);
            }
        }

        private T GetStringValueOfCell<T>(ExcelRange cell)
        {
            var cellValue = cell.Value;

            if (cellValue is DateTime)
            {
                return (T) (object) ((DateTime) cellValue).ToString(dateTimeFormat);
            }

            return (T) (object) cell.Text;
        }

        private void ThrowErrorIfTriedToReadNonExistentWorksheetIsLesserThanZero(int worksheetIndex, int worksheetCount)
        {
            if (worksheetIndex > worksheetCount ||
                worksheetIndex < 0)
                throw new WorksheetReaderNonExistingWorksheetException($"{worksheetIndex}");
        }

        private void ThrowErrorIfRowIsLesserThanZero(int rowIndex)
        {
            if (rowIndex < 1)
                throw new WorksheetReaderRowLesserThanOneException($"{rowIndex}");
        }

        private void ThrowErrorIfColumnIsLesserThanZero(int columnIndex)
        {
            if (columnIndex < 1)
                throw new WorksheetReaderColumnLesserThanOneException($"{columnIndex}");
        }
    }
}