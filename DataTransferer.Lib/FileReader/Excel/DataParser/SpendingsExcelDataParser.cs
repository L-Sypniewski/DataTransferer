using System;
using System.Collections.Generic;
using System.Linq;
using DataTransferer.Lib.ApplicationModel;
using DataTransferer.Lib.ApplicationModel.Excel;
using DataTransferer.Lib.FileReader.DataParser;
using DataTransferer.Lib.FileReader.Excel.WorksheetReader;
using System.Text.RegularExpressions;

namespace DataTransferer.Lib.FileReader.Excel.ExcelDataParser
{
    public class SpendingsExcelDataParser : IDataParser<Spending>
    {
        private readonly IWorksheetReader worksheetReader;
        private const int datesColumnIndex = 1;
        private const int productsColumnIndex = 2;
        private const int categoriesColumnIndex = 3;
        private const int amountColumnIndex = 4;
        private const int firstDateRowIndex = 2;
        private const int maxNumberOfEmptyRowsInBetween = 3;

        public SpendingsExcelDataParser(IWorksheetReader worksheetReader)
        {
            this.worksheetReader = worksheetReader;
        }

        public IEnumerable<Spending> ParseData()
        {
            var parsedSpendings = new List<Spending>();
            var worksheetsCount = worksheetReader.Worksheets.Count();

            for (int worksheetIndex = 0; worksheetIndex < worksheetsCount; worksheetIndex++)
            {
                var firstCellWithDateCoordinates = new ExcelCellCoordinates(worksheetIndex, firstDateRowIndex, datesColumnIndex);
                var rowIndexesContainingDates = GetRowIndexesContainingDates(
                    worksheetReader,
                    firstCellWithDateCoordinates,
                    maxNumberOfEmptyRowsInBetween);


                var lastRowIndexToCheck = rowIndexesContainingDates.Count() != 0 ? rowIndexesContainingDates.Max() + maxNumberOfEmptyRowsInBetween : firstDateRowIndex;
                for (int rowIndex = firstDateRowIndex; rowIndex < lastRowIndexToCheck; rowIndex++)
                {
                    var dateCellCoordinates = new ExcelCellCoordinates(worksheetIndex, rowIndex, datesColumnIndex);
                    var parsedDate = worksheetReader.GetCellDateTimeAsUTC(dateCellCoordinates);

                    if (parsedDate == default(DateTime))
                    {
                        continue;
                    }

                    var productCellCoordinates = new ExcelCellCoordinates(worksheetIndex, rowIndex, productsColumnIndex);
                    var parsedProduct = worksheetReader.GetCellText(productCellCoordinates);

                    var categoryCellCoordinates = new ExcelCellCoordinates(worksheetIndex, rowIndex, categoriesColumnIndex);
                    var parsedCategory = worksheetReader.GetCellText(categoryCellCoordinates);

                    var amountCellCoordinates = new ExcelCellCoordinates(worksheetIndex, rowIndex, amountColumnIndex);
                    var parsedAmount = worksheetReader.GetCellText(amountCellCoordinates);
                    var amountWithoutCurrencyCharacters = Regex.Replace(parsedAmount, "[^0-9.]", "");
                    var convertedAmount = Decimal.Parse(amountWithoutCurrencyCharacters, System.Globalization.NumberStyles.Currency);

                    var parsedCurrency = GetCurrencyFromString(parsedAmount);

                    var parsedSpending = new Spending(parsedDate, parsedProduct, parsedCategory, convertedAmount, parsedCurrency);

                    parsedSpendings.Add(parsedSpending);
                }
            }

            return parsedSpendings;
        }

        private IEnumerable<int> GetRowIndexesContainingDates(
            IWorksheetReader worksheetReader,
            ExcelCellCoordinates firstCellWithDateCoordinates,
            int maxNumberOfEmptyRowsInBetween)
        {
            var firstDatesRowIndex = firstCellWithDateCoordinates.RowIndex;
            var datesColumnIndex = firstCellWithDateCoordinates.ColumnIndex;
            var worksheetIndex = firstCellWithDateCoordinates.WorksheetIndex;

            var rowNumbersContainingDates = new HashSet<int>();
            for (int currentRowIndex = firstDatesRowIndex; ; currentRowIndex++)
            {
                var currentRowCoordinates = new ExcelCellCoordinates(worksheetIndex, currentRowIndex, datesColumnIndex);
                var nextRowsCoordinates = GetCellCoordinatesForNextNRows(currentRowCoordinates, n: maxNumberOfEmptyRowsInBetween);
                var nextRowIndexesContainingDates = GetNextRowIndexesContainingDates(worksheetReader, nextRowsCoordinates);

                if (nextRowIndexesContainingDates.Count() == 0)
                {
                    break;
                }
                rowNumbersContainingDates.UnionWith(nextRowIndexesContainingDates.ToHashSet());
            }

            return rowNumbersContainingDates;
        }

        private Currency GetCurrencyFromString(string cellStringValue)
        {
            return Currency.PLN;
        }

        private IEnumerable<int> GetNextRowIndexesContainingDates(
            IWorksheetReader worksheetReader,
            IEnumerable<ExcelCellCoordinates> nextRowsCoordinatesCollection)
        {
            if (nextRowsCoordinatesCollection.Count() == 0)
            {
                return new int[0];
            }

            var worksheetIndex = nextRowsCoordinatesCollection.First().WorksheetIndex;
            var datesColumnIndex = nextRowsCoordinatesCollection.First().ColumnIndex;

            return nextRowsCoordinatesCollection
                .Select(coordinates => coordinates.RowIndex)
                .Where(rowIndex =>
                {
                    try
                    {
                        var nextRowCoordinates = new ExcelCellCoordinates(worksheetIndex, rowIndex, datesColumnIndex);
                        return worksheetReader.GetCellDateTimeAsUTC(nextRowCoordinates) != default(DateTime);
                    }
                    catch
                    {
                        return false;
                    }
                })
                .ToArray();
        }

        private IEnumerable<ExcelCellCoordinates> GetCellCoordinatesForNextNRows(ExcelCellCoordinates initialCoordinates, int n)
        {
            var firstRowIndex = initialCoordinates.RowIndex;
            var worksheetIndex = initialCoordinates.WorksheetIndex;
            var columnIndex = initialCoordinates.ColumnIndex;

            var excelCellCoordinates = new List<ExcelCellCoordinates>(n);
            for (int rowIndex = firstRowIndex; rowIndex < firstRowIndex + n; rowIndex++)
            {
                var newCellCoordinate = new ExcelCellCoordinates(worksheetIndex, rowIndex, columnIndex);
                excelCellCoordinates.Add(newCellCoordinate);
            }

            return excelCellCoordinates;
        }
    }
}