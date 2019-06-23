using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DataTransferer.Lib.ApplicationModel.Excel;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace DataTransferer.Lib.FileReader.Excel.WorksheetReader
{
    public class WorksheetReader : IWorksheetReader
    {
        private const string dateTimeFormat = "dd-MM-yyyy HH:mm";
        private const string generalExcelCellFormat = "general";
        private readonly CultureInfo cultureInfo = new CultureInfo("en-GB");
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
                        .Select(worksheet => new Worksheet(worksheet.Name))
                        .ToArray();
                }
            }
        }

        public string GetCellText(ExcelCellCoordinates cellCoodrinates)
        {
            using (ExcelPackage package = new ExcelPackage(existingSpreadsheet))
            {
                var worksheets = package.Workbook.Worksheets;
                var cell = GetCellAtCoordinatesFrom(worksheets, cellCoodrinates);

                if (CellNumericFormatEqualsGeneral(cell) == false &&
                CellAtCoordinatesContainsDateTime(cellCoodrinates, new CultureInfo("en-GB")) == false)
                {
                    return cell.Text;
                }

                if (CellAtCoordinatesContainsDateTime(cellCoodrinates, cultureInfo))
                {
                    var dateTimeUTC = GetCellDateTime(cellCoodrinates).ToUniversalTime();
                    return dateTimeUTC.ToString("yyyy-MM-ddThh:mm:ssZ");
                }

                try
                {
                    return cell.GetValue<string>() ?? "";
                }
                catch { }

                return cell.Text;
            }
        }

        public bool CellAtCoordinatesContainsDateTime(ExcelCellCoordinates cellCoordinates, CultureInfo cultureInfo)
        {
            using (ExcelPackage package = new ExcelPackage(existingSpreadsheet))
            {
                int worksheetIndex = cellCoordinates.WorksheetIndex;
                int rowIndex = cellCoordinates.RowIndex;
                int columnIndex = cellCoordinates.ColumnIndex;
                var cell = package.Workbook.Worksheets[worksheetIndex].Cells[rowIndex, columnIndex];

                var cellStringValue = cell.GetValue<string>();
                var cellText = cell.Text;

                if (cellStringValue == null && cellText == "")
                {
                    return false;
                }

                cellStringValue = cellStringValue.Replace(".", "");
                cellText = cellText.Replace(".", "");

                try
                {
                    var cellDateTimeValue = cell.GetValue<DateTime>();

                    if (cellStringValue == cellText)
                    {
                        return true;
                    }

                    var isCellStringValueOfDateTimeType = DateTime.TryParse(
                        cellStringValue,
                        cultureInfo,
                        DateTimeStyles.AssumeLocal,
                        out DateTime datetimeStringValueParsedData);

                    var isCellTextOfDateTimeType = DateTime.TryParse(
                        cellText,
                        cultureInfo,
                        DateTimeStyles.AssumeLocal,
                        out DateTime datetimeTextParsedData);


                    bool cellValueAndCellTextAreNotDateTimeType = isCellStringValueOfDateTimeType == false &&
                                                                  isCellTextOfDateTimeType == false;
                    if (cellValueAndCellTextAreNotDateTimeType)
                    {
                        return false;
                    }

                    if ((cellStringValue.All(char.IsDigit) && datetimeTextParsedData == cellDateTimeValue) ||
                        bothDatesHaveSameYearMonthAndDay(datetimeTextParsedData, cellDateTimeValue))
                    {
                        return true;
                    }

                    return false;
                }
                catch
                {
                    return false;
                }
            }
        }

        private bool bothDatesHaveSameYearMonthAndDay(DateTime date1, DateTime date2)
        {
            return date1.Year == date2.Year &&
                   date1.Month == date2.Month &&
                   date1.Day == date2.Day;
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
            using (ExcelPackage package = new ExcelPackage(existingSpreadsheet))
            {
                var worksheets = package.Workbook.Worksheets;
                var cell = GetCellAtCoordinatesFrom(worksheets, cellCoodrinates);

                if (CellAtCoordinatesContainsDateTime(cellCoodrinates, cultureInfo) == false)
                {
                    throw new WorksheetReaderCellValueTypeException(cell.Text, cellCoodrinates, typeof(DateTime));
                }

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