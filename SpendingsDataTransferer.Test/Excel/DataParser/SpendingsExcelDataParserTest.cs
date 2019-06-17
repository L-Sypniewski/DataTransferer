using System;
using System.Linq;
using FluentAssertions;
using SpendingsDataTransferer.Lib.ApplicationModel;
using SpendingsDataTransferer.Lib.ApplicationModel.Excel;
using SpendingsDataTransferer.Lib.FileReader.DataParser;
using SpendingsDataTransferer.Lib.FileReader.Excel.ExcelDataParser;
using SpendingsDataTransferer.Lib.FileReader.Excel.WorksheetReader;
using Xunit;

namespace SpendingsDataTransferer.Test
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
            var expectedSpendingsCount = 21;

            spednings
            .Should()
            .HaveCount(expectedSpendingsCount);
        }
    }
}