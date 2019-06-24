using FluentAssertions;
using DataTransferer.Lib.FileReader.Excel.WorksheetReader;
using Xunit;
using Moq;
using DataTransferer.Lib.ApplicationModel.Excel;
using System;
using DataTransferer.Lib.FileReader.Excel.ExcelDataParser;

namespace DataTransferer.Test
{
    public class SpendingsExcelDataParserTest
    {
        const string filepath = @"TestFiles/testFile.xlsx";
        readonly Worksheet[] worksheets = { new Worksheet(), new Worksheet(), new Worksheet(), new Worksheet() };


        [Fact]
        public void CheckNumberOfSpendingsThatContainDatesInARowAllInOneWorksheet()
        {
            var worksheetReaderMock = new Mock<IWorksheetReader>();

            worksheetReaderMock.Setup(worksheetReader => worksheetReader.Worksheets).Returns(worksheets);

            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 2, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 3, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 4, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 5, 1))).Returns(DateTime.Now);

            var excelSpendingDataParser = new SpendingsExcelDataParser(worksheetReaderMock.Object);


            var spednings = excelSpendingDataParser.ParseData();
            var expectedSpendingsCount = 4;


            spednings.Should().HaveCount(expectedSpendingsCount);
        }

        [Fact]
        public void CheckNumberOfSpendingsThatContainDatesInARowInThreeWorksheets()
        {
            var worksheetReaderMock = new Mock<IWorksheetReader>();

            worksheetReaderMock.Setup(worksheetReader => worksheetReader.Worksheets).Returns(worksheets);

            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 2, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 3, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 4, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(1, 2, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(1, 3, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(2, 2, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(2, 3, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(2, 4, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(2, 5, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(2, 6, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(3, 2, 1))).Returns(DateTime.Now);

            var excelSpendingDataParser = new SpendingsExcelDataParser(worksheetReaderMock.Object);


            var spednings = excelSpendingDataParser.ParseData();
            var expectedSpendingsCount = 11;


            spednings.Should().HaveCount(expectedSpendingsCount);
        }

        [Fact]
        public void CheckNumberOfSpendingsThatContainDatesWithSpacesBetweenInOneWorksheet()
        {
            var worksheetReaderMock = new Mock<IWorksheetReader>();

            worksheetReaderMock.Setup(worksheetReader => worksheetReader.Worksheets).Returns(worksheets);

            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 2, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 3, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 6, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 8, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 9, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 14, 1))).Returns(DateTime.Now);
            worksheetReaderMock.Setup(worksheetReader => worksheetReader.GetCellDateTimeAsUTC(new ExcelCellCoordinates(0, 15, 1))).Returns(DateTime.Now);

            var excelSpendingDataParser = new SpendingsExcelDataParser(worksheetReaderMock.Object);


            var spednings = excelSpendingDataParser.ParseData();
            var expectedSpendingsCount = 5;


            spednings.Should().HaveCount(expectedSpendingsCount);
        }
    }
}