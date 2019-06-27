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
using DataTransferer.Lib.ApplicationModel;

namespace DataTransferer.Test
{
    using RowsAndColumnsIndexes = List<(int worksheetIndex, int rowIndex)>;
    using CoordinatesWithDataList = List<(ExcelCellCoordinates cellCoordinates, object cellData)>;
    using Spendings = List<Spending>;

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
        public void ParsingValidSpendingsTest(CoordinatesWithDataList coordinatesWithData, Spendings expectedSpendings)
        {
            var worksheetReaderMock = WorksheetReaderMockWithValidSpendingsDataAt(coordinatesWithData);
            var excelSpendingDataParser = new SpendingsExcelDataParser(worksheetReaderMock);

            var sortedSpendings = excelSpendingDataParser.ParseData().ToList();
            sortedSpendings.Sort();

            var sortedExpectedSpendings = expectedSpendings;
            sortedExpectedSpendings.Sort();

            sortedSpendings.Should().Equal(excelSpendingDataParser);
        }

        private IWorksheetReader WorksheetReaderMockWithValidSpendingsDataAt(CoordinatesWithDataList coordinatesWithData)
        {
            var worksheetReaderMock = new Mock<IWorksheetReader>();

            var numberOfWorksheets = coordinatesWithData
                .Select(x => x.cellCoordinates.WorksheetIndex)
                .Max() + 1;
            Worksheet[] worksheets = Enumerable.Repeat(new Worksheet(), numberOfWorksheets).ToArray();

            worksheetReaderMock.Setup(worksheetReader => worksheetReader.Worksheets).Returns(worksheets);


            foreach (var dataAtCoordinates in coordinatesWithData)
            {
                var cellData = dataAtCoordinates.cellData;
                var cellCoordinates = dataAtCoordinates.cellCoordinates;
                if (cellData is string stringData)
                {
                    worksheetReaderMock.Setup(worksheetReader =>
                      worksheetReader.GetCellText(cellCoordinates)).Returns(stringData);
                }
                else if (cellData is DateTime dateTimeData)
                {
                    worksheetReaderMock.Setup(worksheetReader =>
                     worksheetReader.GetCellDateTimeAsUTC(cellCoordinates)).Returns(dateTimeData);
                }
                else
                {
                    throw new ArgumentException($"Unsupported type for worksheetReader to read: type equals {cellData.GetType()}");
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
                var coordinatesWithDataList = new CoordinatesWithDataList {
                    (new ExcelCellCoordinates(1, 3, 1), new DateTime(2017, 12, 4)),
                    (new ExcelCellCoordinates(1, 3, 2), "Coca Cola"),
                    (new ExcelCellCoordinates(1, 3, 3), "Jedzenie"),
                    (new ExcelCellCoordinates(1, 3, 4), "5.99 zł"),
                };

                var expectedParsedSpendings = new Spendings
                {
                    new Spending(new DateTime(2017, 12, 4), "Coca Cola", "Jedzenie", 5.99m, Currency.PLN),
                };


                yield return new object[] { coordinatesWithDataList, expectedParsedSpendings };
            }
        }
    }
}