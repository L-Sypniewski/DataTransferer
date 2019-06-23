using System;
using System.Collections.Generic;
using System.Linq;
using DataTransferer.Lib.ApplicationModel;
using DataTransferer.Lib.ApplicationModel.Excel;
using DataTransferer.Lib.FileReader.DataParser;
using DataTransferer.Lib.FileReader.Excel.WorksheetReader;

namespace DataTransferer.Lib.FileReader.Excel.ExcelDataParser
{
    public class SpendingsExcelDataParser : IDataParser<Spending>
    {
        private readonly IWorksheetReader worksheetReader;
        private const int datesColumnIndex = 1;
        private const int firstDateRowIndex = 2;
        private const int maxNumberOfEmptyRowsInBetween = 3;

        public SpendingsExcelDataParser(IWorksheetReader worksheetReader)
        {
            this.worksheetReader = worksheetReader;
        }

        public IEnumerable<Spending> ParseData()
        {
            var datesRowNumbers = new List<Spending>();
            var worksheetsCount = worksheetReader.Worksheets.Count();

            for (int worksheetIndex = 0; worksheetIndex < worksheetsCount; worksheetIndex++)
            {
                var firstDateCellCoordinates = new ExcelCellCoordinates(worksheetIndex, firstDateRowIndex, datesColumnIndex);
                var rowIndexesContainingDates = GetRowIndexesContainingDates(
                    worksheetReader,
                    firstDateCellCoordinates,
                    maxNumberOfEmptyRowsInBetween);

                for (int rowIndex = firstDateRowIndex; rowIndex < firstDateRowIndex + rowIndexesContainingDates.Count(); rowIndex++)
                {
                    datesRowNumbers.Add(new Spending());
                }
            }

            return datesRowNumbers;
        }

        private IEnumerable<int> GetRowIndexesContainingDates(
            IWorksheetReader worksheetReader,
            ExcelCellCoordinates firstDateCellCoordinates,
            int maxNumberOfEmptyRowsInBetween)
        {
            var firstDatesRowIndex = firstDateCellCoordinates.RowIndex;
            var datesColumnIndex = firstDateCellCoordinates.ColumnIndex;
            var worksheetIndex = firstDateCellCoordinates.WorksheetIndex;

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