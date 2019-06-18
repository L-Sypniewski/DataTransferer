using System;
using System.Linq;
using DataTransferer.Lib.ApplicationModel.Excel;
using DataTransferer.Lib.FileReader.Excel.WorksheetReader;
using FluentAssertions;
using Xunit;

namespace DataTransferer.Test
{
    public class ExcelReadingTest
    {
        const string filepath = @"TestFiles/testFile.xlsx";

        IWorksheetReader worksheetReader = new WorksheetReader(filepath);

        [Fact]
        public void CheckNumberOfWorksheets()
        {
            var worksheets = worksheetReader.Worksheets.ToList();
            var worksheetsCount = worksheets.Count();
            var expectedNumberOfWorksheets = 3;

            worksheets
                .Should()
                .HaveCount(expectedNumberOfWorksheets, $"the file contains {expectedNumberOfWorksheets} worksheets");
        }

        [Theory]
        [InlineData(0)]
        [InlineData(-1)]
        [InlineData(-10)]
        public void TryingToGetCellTextFromIncorrectRowIndexThrowsError(int rowIndex)
        {
            var coordinates = new ExcelCellCoordinates(worksheetIndex: 1, rowIndex: rowIndex, columnIndex: 1);

            Action act = () => worksheetReader.GetCellText(coordinates);
            act
            .Should()
            .ThrowExactly<WorksheetReaderRowLesserThanOneException>("row index must be larger than or equal 1")
            .WithMessage($"WorksheetReader tried to read row at index: {rowIndex}");

        }

        [Theory]
        [InlineData(0, 1, 1, "Data")]
        [InlineData(1, 4, 1, "05-12-2017 00:00")]
        [InlineData(1, 19, 1, "30-12-2017 00:00")]
        [InlineData(1, 20, 1, "")]
        [InlineData(0, 3, 4, "119.95 zł")]
        [InlineData(0, 8, 5, "")]
        [InlineData(2, 5, 2, "Bilet miesięczny")]
        [InlineData(2, 11, 3, "Fast food")]
        public void CellReadingStringDataTest(int worksheetIndex, int rowIndex, int columnIndex, string expectedCellValue)
        {
            var celCoordinates = new ExcelCellCoordinates(worksheetIndex, rowIndex, columnIndex);
            var cellValue = worksheetReader.GetCellText(celCoordinates);
            cellValue
                .Should()
                .Be(expectedCellValue);
        }
    }
}