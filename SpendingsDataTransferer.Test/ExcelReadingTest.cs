using System.Linq;
using FluentAssertions;
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
    }
}