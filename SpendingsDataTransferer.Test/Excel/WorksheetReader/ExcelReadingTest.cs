using System;
using System.Linq;
using FluentAssertions;
using SpendingsDataTransferer.Lib.ApplicationModel.Excel;
using SpendingsDataTransferer.Lib.FileReader.Excel.WorksheetReader;
using Xunit;

namespace SpendingsDataTransferer.Test
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
        [InlineData(0, 4, 1, "10-25-2017 00:00")]
        [InlineData(0, 10, 4, "18.78 zł")]
        [InlineData(0, 1, 5, "")]
        [InlineData(2, 8, 2, "Pizza Tornado")]
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