using System;
using System.Linq;
using FluentAssertions;
using DataTransferer.Lib.ApplicationModel;
using DataTransferer.Lib.ApplicationModel.Excel;
using DataTransferer.Lib.FileReader.DataParser;
using DataTransferer.Lib.FileReader.Excel.ExcelDataParser;
using DataTransferer.Lib.FileReader.Excel.WorksheetReader;
using Xunit;

namespace DataTransferer.Test
{
    public class SpendingsExcelDataParserTest
    {
        const string filepath = @"TestFiles/testFile.xlsx";

        static IWorksheetReader worksheetReader = new WorksheetReader(filepath);
        static IDataParser<Spending> excelSpendingDataParser = new SpendingsExcelDataParser(worksheetReader);


        [Fact]
        public void CheckNumberOfSpendings()
        {
            var spednings = excelSpendingDataParser.ParseData();
            var expectedSpendingsCount = 20;

            spednings
            .Should()
            .HaveCount(expectedSpendingsCount);
        }
    }
}