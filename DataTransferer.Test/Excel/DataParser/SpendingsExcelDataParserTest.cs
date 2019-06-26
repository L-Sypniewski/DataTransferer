using FluentAssertions;
using DataTransferer.Lib.FileReader.Excel.WorksheetReader;
using Xunit;
using Moq;
using DataTransferer.Lib.ApplicationModel.Excel;
using System;
using DataTransferer.Lib.FileReader.Excel.ExcelDataParser;
using System.Collections.Generic;
using System.Collections;
using System.Linq;

namespace DataTransferer.Test
{
    using RowsAndColumnsIndexes = List<(int worksheetIndex, int rowIndex)>;
    using CoordinatesWithDataEnumerable = IEnumerable<(ExcelCellCoordinates cellCoordinates, object objectAtCell)>;

    public class SpendingsExcelDataParserTest
    {
        const string filepath = @"TestFiles/testFile.xlsx";
        const int dateTimeColumnIndex = 1;

        [Theory]
        [ClassData(typeof(CheckNumberOfSpendingsData))]
        public void CheckNumberOfSpendings(string @if, RowsAndColumnsIndexes indexesOfCellsWithValidDates, int expectedSpedningsNumber)
        {
            var worksheetReaderMock = WorksheetReaderMockWithValidDateTimesAt(indexesOfCellsWithValidDates);
            var excelSpendingDataParser = new SpendingsExcelDataParser(worksheetReaderMock);

            var spednings = excelSpendingDataParser.ParseData();

            spednings.Should().HaveCount(expectedSpedningsNumber, $"that's the expected number of valid spednings if {@if}");
        }

        public class CheckNumberOfSpendingsData : IEnumerable<object[]>
        {
            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public IEnumerator<object[]> GetEnumerator()
            {
                yield return new object[] { "all cells contanining dates are in row in 1 worksheet",
                    new RowsAndColumnsIndexes { (0, 2), (0, 3), (0, 4), (0, 5) }, 4 };

                yield return new object[] { "all cells contanining dates are in row in 4 worksheets",
                    new RowsAndColumnsIndexes { (0, 2), (0, 3), (0, 4), (1, 2), (1, 3), (2, 2),
                                                               (2, 3), (2, 4), (2, 5), (2, 6), (3, 2), }, 11 };

                yield return new object[] { "all cells contanining dates are with spaces in between in 1 worksheets",
                    new RowsAndColumnsIndexes { (0, 2), (0, 3), (0, 6), (0, 8), (0, 9), (0, 14), (0, 15) }, 5 };
            }
        }

        private IWorksheetReader WorksheetReaderMockWithValidDateTimesAt(RowsAndColumnsIndexes coordinatesList)
        {
            var worksheetReaderMock = new Mock<IWorksheetReader>();
            Worksheet[] worksheets = Enumerable.Repeat(new Worksheet(), 4).ToArray();

            worksheetReaderMock.Setup(worksheetReader => worksheetReader.Worksheets).Returns(worksheets);

            foreach (var indexes in coordinatesList)
            {
                var cellCoordinates = new ExcelCellCoordinates(indexes.worksheetIndex, indexes.rowIndex, dateTimeColumnIndex);
                worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(cellCoordinates)).Returns(DateTime.Now);
            }

            return worksheetReaderMock.Object;
        }


        [Theory]
        [ClassData(typeof(ParsingValidSpendingsData))]
        public void ParsingValidSpendingsTest(CoordinatesWithDataEnumerable coordinatesWithData)
        {
            var worksheetReaderMock = WorksheetReaderMockWithValidSpendingsDataAt(coordinatesWithData);
            var excelSpendingDataParser = new SpendingsExcelDataParser(worksheetReaderMock);

            var spednings = excelSpendingDataParser.ParseData();


        }

        private IWorksheetReader WorksheetReaderMockWithValidSpendingsDataAt(CoordinatesWithDataEnumerable coordinatesWithData)
        {
            var worksheetReaderMock = new Mock<IWorksheetReader>();

            var numberOfWorksheets = coordinatesWithData
                .Select(x => x.cellCoordinates.WorksheetIndex)
                .Max();
            Worksheet[] worksheets = Enumerable.Repeat(new Worksheet(), numberOfWorksheets).ToArray();


            foreach (var dataAtCoordinates in coordinatesWithData)
            {
                var objectAtCell = dataAtCoordinates.objectAtCell;
                var cellCoordinates = dataAtCoordinates.cellCoordinates;
                if (objectAtCell is string stringData)
                {
                    worksheetReaderMock.Setup(worksheetReader =>
                      worksheetReader.GetCellText(cellCoordinates)).Returns(stringData);
                }
                else if (objectAtCell is DateTime dateTimeData)
                {
                    worksheetReaderMock.Setup(worksheetReader =>
                     worksheetReader.GetCellDateTimeAsUTC(cellCoordinates)).Returns(dateTimeData);
                }
                else
                {
                    throw new ArgumentException($"Unsupported type for worksheetReader to read: type equals {objectAtCell.GetType()}");
                }
            }

            return worksheetReaderMock.Object;
        }


        public class ParsingValidSpendingsData : IEnumerable<object[]>
        {
            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public IEnumerator<object[]> GetEnumerator()
            {
                yield return new object[] { "all cells contanining dates are in row in 1 worksheet",
                    new (int worksheetIndex, int rowIndex)[] { (0, 2), (0, 3), (0, 4), (0, 5) }, 4 };

                yield return new object[] { "all cells contanining dates are in row in 4 worksheets",
                    new (int worksheetIndex, int rowIndex)[] { (0, 2), (0, 3), (0, 4), (1, 2), (1, 3), (2, 2),
                                                               (2, 3), (2, 4), (2, 5), (2, 6), (3, 2), }, 11 };

                yield return new object[] { "all cells contanining dates are with spaces in between in 1 worksheets",
                    new (int worksheetIndex, int rowIndex)[] { (0, 2), (0, 3), (0, 6), (0, 8), (0, 9), (0, 14), (0, 15) }, 5 };
            }
        }
    }
}