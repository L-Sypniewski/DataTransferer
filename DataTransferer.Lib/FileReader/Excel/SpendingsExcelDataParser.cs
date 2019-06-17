using System.Collections.Generic;
using DataTransferer.Lib.ApplicationModel;
using DataTransferer.Lib.ApplicationModel.Excel;
using DataTransferer.Lib.FileReader.DataParser;
using DataTransferer.Lib.FileReader.Excel.WorksheetReader;
using System.Linq;
using System;

namespace DataTransferer.Lib.FileReader.Excel.ExcelDataParser
{
    public class SpendingsExcelDataParser : IDataParser<Spending>
    {
        private readonly string filename;
        private readonly IWorksheetReader worksheetReader;
        private const int indexOfColumnWithDates = 1;
        private const int firstDatesRowIndex = 2;
        private const int maxNumberOfEmptyRowsInBetween = 3;


        public SpendingsExcelDataParser(IWorksheetReader worksheetReader)
        {
            this.worksheetReader = worksheetReader;
        }

        public IEnumerable<Spending> ParseData()
        {
            var rowNumbersWithDates = new List<Spending>();
            var worksheetsCount = worksheetReader.Worksheets.Count();

            for (int worksheetIndex = 0; worksheetIndex < worksheetsCount; worksheetIndex++)
            {
                var rowNumbersContainingDatesInWorksheet = GetRowNumbersContainingDates(
                    worksheetReader, worksheetIndex, indexOfColumnWithDates, firstDatesRowIndex, maxNumberOfEmptyRowsInBetween);

                for (int rowIndex = firstDatesRowIndex; rowIndex < firstDatesRowIndex + rowNumbersContainingDatesInWorksheet.Count(); rowIndex++)
                {
                    rowNumbersWithDates.Add(new Spending());
                }
            }

            return rowNumbersWithDates;
        }

        private IEnumerable<int> GetRowNumbersContainingDates(IWorksheetReader worksheetReader, int worksheetIndex,
            int datesColumnIndex, int firstDatesRowIndex, int maxNumberOfEmptyRowsInBetween)
        {
            var rowNumbersContainingDates = new HashSet<int>();
            for (int currentRowIndex = firstDatesRowIndex; ; currentRowIndex++)
            {
                System.Console.WriteLine($"Checking row number {currentRowIndex}");
                var nextCellsCoordinates = GetCellCoordinatesForNextNRowsInColumn(
                    worksheetIndex, currentRowIndex, maxNumberOfEmptyRowsInBetween, datesColumnIndex);

                var nextDates = nextCellsCoordinates
                .Select(coordinates => worksheetReader.GetCellDateTime(coordinates));

                var nextRowNumbers = nextCellsCoordinates
                .Select(coordinates => coordinates.RowIndex)
                .Where(rowIndex =>
                {
                    try
                    {
                        return worksheetReader.GetCellDateTime(
                            new ExcelCellCoordinates(worksheetIndex, rowIndex, datesColumnIndex)) != default(DateTime);
                    }
                    catch
                    {
                        return false;
                    }
                })
                    .ToArray();

                if (nextRowNumbers.Length == 0)
                {
                    break;
                }
                rowNumbersContainingDates.UnionWith(nextRowNumbers.ToHashSet());
            }
            return rowNumbersContainingDates;
        }

        private IEnumerable<ExcelCellCoordinates> GetCellCoordinatesForNextNRowsInColumn(
            int worksheetIndex, int firstRowIndex, int n, int columnIndex)
        {
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