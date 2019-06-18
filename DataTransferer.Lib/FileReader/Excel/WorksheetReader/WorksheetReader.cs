using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DataTransferer.Lib.ApplicationModel.Excel;
using OfficeOpenXml;

namespace DataTransferer.Lib.FileReader.Excel.WorksheetReader
{
    public class WorksheetReader : IWorksheetReader
    {
        private const string dateTimeFormat = "dd-MM-yyyy HH:mm";
        private const string generalExcelCellFormat = "general";
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

                try
                {
                    return cell.GetValue<string>();
                }
                catch { }

                return cell.Text;
            }
        }

        public bool IsCellContainDateTime(ExcelCellCoordinates coordinates)
        {
            using(ExcelPackage package = new ExcelPackage(existingSpreadsheet))
            {
                var cell = package.Workbook.Worksheets[coordinates.WorksheetIndex].Cells[coordinates.RowIndex, coordinates.ColumnIndex];

                var datetimeString = cell.GetValue<string>();

                var regex = new RegEgc
                var datetime = cell.GetValue<DateTime>();

                var isCellContainDateTime = DateTime.TryParse(datetimeString, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime date);

                return isCellContainDateTime;
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

        private bool CellNumericFormatEqualsGeneral(ExcelRange cell)
        {
            return cell.Style.Numberformat.Format == generalExcelCellFormat;
        }

        public DateTime GetCellDateTime(ExcelCellCoordinates cellCoodrinates)
        {
            using(ExcelPackage package = new ExcelPackage(existingSpreadsheet))
            {
                var worksheets = package.Workbook.Worksheets;
                var cell = GetCellAtCoordinatesFrom(worksheets, cellCoodrinates);

                try
                {
                    return cell.GetValue<DateTime>();
                }
                catch (FormatException e)
                {
                    throw new WorksheetReaderCellValueTypeException(cell.Text, cellCoodrinates, typeof(DateTime), e);
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
            if (worksheetIndex >= worksheetCount ||
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