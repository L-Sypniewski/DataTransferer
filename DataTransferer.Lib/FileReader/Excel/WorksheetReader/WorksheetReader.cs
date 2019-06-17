using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using DataTransferer.Lib.ApplicationModel.Excel;

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
                using (ExcelPackage package = new ExcelPackage(existingSpreadsheet))
                {
                    return package.Workbook.Worksheets
                        .Select(worksheet => new Worksheet(worksheet.Name)).ToArray();
                }
            }
        }

        public string GetCellText(ExcelCellCoordinates cellCoodrinates)
        {
            using (ExcelPackage package = new ExcelPackage(existingSpreadsheet))
            {
                var worksheets = package.Workbook.Worksheets;

                ThrowErrorIfCoorinatesAreIncorrect(cellCoodrinates, worksheets);

                var worksheetIndex = cellCoodrinates.WorksheetIndex;
                var rowIndex = cellCoodrinates.RowIndex;
                var columnIndex = cellCoodrinates.ColumnIndex;

                var cell = worksheets[worksheetIndex].Cells[rowIndex, columnIndex];

                if (cell.Style.Numberformat.Format != "general")
                {
                    return cell.Text;
                }

                try
                {
                    return GetDateTimeAsTypeFromCell<String>(worksheets, cellCoodrinates,
                                                             cellValueAsDateTime => cellValueAsDateTime.ToString(dateTimeFormat));
                }
                catch { }

                return cell.Text;
            }
        }

        private T GetDateTimeAsTypeFromCell<T>(ExcelWorksheets worksheets,
                                               ExcelCellCoordinates cellCoordinates,
                                               Func<DateTime, T> funcReturningDateTimeAsType)
        {
            var cellValueAsDateTime = worksheets[cellCoordinates.WorksheetIndex].GetValue<DateTime>(
                cellCoordinates.RowIndex, cellCoordinates.ColumnIndex);

            if (cellValueAsDateTime != default(DateTime))
            {
                return funcReturningDateTimeAsType(cellValueAsDateTime);
            }

            throw new WorksheetReaderCellValueTypeException(cellCoordinates, typeof(DateTime));
        }

        public DateTime GetCellDateTime(ExcelCellCoordinates cellCoodrinates)
        {
            using (ExcelPackage package = new ExcelPackage(existingSpreadsheet))
            {
                var worksheets = package.Workbook.Worksheets;

                ThrowErrorIfCoorinatesAreIncorrect(cellCoodrinates, worksheets);

                var worksheetIndex = cellCoodrinates.WorksheetIndex;
                var rowIndex = cellCoodrinates.RowIndex;
                var columnIndex = cellCoodrinates.ColumnIndex;

                try
                {
                    return GetDateTimeAsTypeFromCell<DateTime>(worksheets, cellCoodrinates,
                                                     cellValueAsDateTime => cellValueAsDateTime);
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
            ThrowErrorIfTriedToReadNonExistentWorksheetIsLesserThanZero(cellCoodrinates.WorksheetIndex, worksheets.Count);
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

        private void ThrowExceptionIfDateTimeWasMappedToDefaultValue(DateTime cellValue, ExcelCellCoordinates cellCoodrinates)
        {
            if (cellValue == default(DateTime))
            {
                throw new WorksheetReaderCellValueTypeException(cellCoodrinates, typeof(DateTime));
            }
        }
    }
}