using System;
using System.Linq;
using FluentAssertions;
using DataTransferer.Lib.ApplicationModel.Excel;
using DataTransferer.Lib.FileReader.Excel.WorksheetReader;
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
            var expectedNumberOfWorksheets = 21;

            worksheets
                .Should()
                .HaveCount(expectedNumberOfWorksheets, $"the file contains {expectedNumberOfWorksheets} worksheets");
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