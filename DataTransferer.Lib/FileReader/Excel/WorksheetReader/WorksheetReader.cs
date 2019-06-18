using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DataTransferer.Lib.ApplicationModel.Excel;
using OfficeOpenXml;

namespace DataTransferer.Lib.FileReader.Excel.WorksheetReader
{
    public class WorksheetReader : IWorksheetReader
    {
        private const string dateTimeFormat = "dd-MM-yyyy HH:mm";
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
                        .Select(worksheet => new Worksheet(worksheet.Name))
                        .ToArray();
                }
            }
        }

        public string GetCellText(ExcelCellCoordinates cellCoodrinates)
        {
            using(ExcelPackage package = new ExcelPackage(existingSpreadsheet))
            {
                var worksheets = package.Workbook.Worksheets;
                var cell = GetCellAtCoordinatesFrom(worksheets, cellCoodrinates);

                if (cell.Style.Numberformat.Format != "general")
                {
                    return cell.Text;
                }

                try
                {
                    return GetDateTimeAsTypeFromCell<String>(cell, cellValueAsDateTime => cellValueAsDateTime.ToString(dateTimeFormat));
                }
                catch { }

                return cell.Text;
            }
        }

        private ExcelRange GetCellAtCoordinatesFrom(ExcelWorksheets worksheets, ExcelCellCoordinates cellCoodrinates)
        {
            ThrowErrorIfCoorinatesAreIncorrect(cellCoodrinates, worksheets);

            var worksheetIndex = cellCoodrinates.WorksheetIndex;
            var rowIndex = cellCoodrinates.RowIndex;
            var columnIndex = cellCoodrinates.ColumnIndex;

            var cellAtCoordinates = worksheets[worksheetIndex].Cells[rowIndex, columnIndex];
            return cellAtCoordinates;
        }

        private T GetDateTimeAsTypeFromCell<T>(ExcelRange cell, Func<DateTime, T> funcReturningDateTimeAsType)
        {
            var cellValueAsDateTime = cell.GetValue<DateTime>();

            if (cellValueAsDateTime != default(DateTime))
            {
                return funcReturningDateTimeAsType(cellValueAsDateTime);
            }
            var cellCoordinates = new ExcelCellCoordinates(cell.Worksheet.Index, cell.Rows, cell.Columns);
            throw new WorksheetReaderCellValueTypeException(cellCoordinates, typeof(DateTime));
        }

        public DateTime GetCellDateTime(ExcelCellCoordinates cellCoodrinates)
        {
            using(ExcelPackage package = new ExcelPackage(existingSpreadsheet))
            {
                var worksheets = package.Workbook.Worksheets;
                var cell = GetCellAtCoordinatesFrom(worksheets, cellCoodrinates);

                try
                {
                    return GetDateTimeAsTypeFromCell<DateTime>(cell, cellValueAsDateTime => cellValueAsDateTime);
                }
                catch (FormatException e)
                {
                    throw new WorksheetReaderCellValueTypeException(cellCoodrinates, typeof(DateTime), e);
                }
            }
        }

        private void ThrowErrorIfCoorinatesAreIncorrect(ExcelCellCoordinates cellCoodrinates, ExcelWorksheets worksheets)
        {
            ThrowErrorIfRowIsLesserThanZero(cellCoodrinates.RowIndex);
            ThrowErrorIfColumnIsLesserThanZero(cellCoodrinates.ColumnIndex);
            ThrowErrorIfTriedToReadNonExistentWorksheet(cellCoodrinates.WorksheetIndex, worksheets.Count);
        }

        private void ThrowErrorIfTriedToReadNonExistentWorksheet(int worksheetIndex, int worksheetCount)
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

        private void ThrowExceptionIfDateTimeWasMappedToDefaultValue(DateTime cellValue, ExcelCellCoordinates cellCoodrinates)
        {
            if (cellValue == default(DateTime))
            {
                throw new WorksheetReaderCellValueTypeException(cellCoodrinates, typeof(DateTime));
            }
        }
    }
}