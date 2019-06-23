using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DataTransferer.Lib.ApplicationModel.Excel;
using DataTransferer.Lib.FileReader.Excel.WorksheetReader;
using FluentAssertions;
using Xunit;

namespace DataTransferer.Test
{
    public class WorksheetReaderTest
    {
        const string filepath = @"TestFiles/testFile.xlsx";

        [Fact]
        public void CheckNumberOfWorksheets()
        {
            IWorksheetReader worksheetReader = new WorksheetReader(filepath);

            var worksheets = worksheetReader.Worksheets.ToList();
            var worksheetsCount = worksheets.Count();
            var expectedNumberOfWorksheets = 3;

            worksheets
                .Should()
                .HaveCount(expectedNumberOfWorksheets, $"the file contains {expectedNumberOfWorksheets} worksheets");
        }

        [Theory]
        [ClassData(typeof(GettingCellDataFromIncorrectRowOrColumnIndexesData))]
        public void TryingToGetCellDataFromIncorrectRowIndexThrowsError(int rowIndex, Action<ExcelCellCoordinates> getCellDataAction)
        {
            IWorksheetReader worksheetReader = new WorksheetReader(filepath);

            var coordinates = new ExcelCellCoordinates(worksheetIndex: 1, rowIndex: rowIndex, columnIndex: 1);

            Action getCellTextAction = () => getCellDataAction(coordinates);
            getCellTextAction
                .Should()
                .ThrowExactly<WorksheetReaderRowLesserThanOneException>("row index must be larger than or equal 1")
                .WithMessage($"WorksheetReader tried to read row at index: {rowIndex}");

        }

        [Theory]
        [ClassData(typeof(GettingCellDataFromIncorrectRowOrColumnIndexesData))]
        public void TryingToGetCellDataFromIncorrectColumnIndexThrowsError(int columnIndex, Action<ExcelCellCoordinates> getCellDataAction)
        {
            IWorksheetReader worksheetReader = new WorksheetReader(filepath);

            var coordinates = new ExcelCellCoordinates(worksheetIndex: 1, rowIndex: 1, columnIndex: columnIndex);

            Action getCellTextAction = () => getCellDataAction(coordinates);
            getCellTextAction
                .Should()
                .ThrowExactly<WorksheetReaderColumnLesserThanOneException>("column index must be larger than or equal 1")
                .WithMessage($"WorksheetReader tried to read column at index: {columnIndex}");
        }

        [Theory]
        [ClassData(typeof(GettingCellDataFromNonExistingWorksheetData))]
        public void TryingToGetCellDataFromNonExistingWorksheetThrowsError(int worksheetIndex, Action<ExcelCellCoordinates> getCellDataAction)
        {
            IWorksheetReader worksheetReader = new WorksheetReader(filepath);

            var coordinates = new ExcelCellCoordinates(worksheetIndex: worksheetIndex, rowIndex: 1, columnIndex: 1);

            Action getCellTextAction = () => getCellDataAction(coordinates);
            getCellTextAction
                .Should()
                .ThrowExactly<WorksheetReaderNonExistingWorksheetException>("worksheet index cannot be less than zero or higher than worksheets count")
                .WithMessage($"There is no worksheet at index: {worksheetIndex}");
        }

        [Theory]
        [InlineData(0, 1, 1, "Data")]
        [InlineData(1, 4, 1, "2017-12-05T00:00:00Z")]
        [InlineData(1, 19, 1, "2017-12-30T00:00:00Z")]
        [InlineData(0, 17, 1, "2012-03-08T06:15:00Z")]
        [InlineData(0, 19, 1, "1999-06-12T13:46:27Z")]
        [InlineData(1, 20, 1, "")]
        [InlineData(0, 8, 5, "")]
        [InlineData(0, 3, 4, "119.95 zł")]
        [InlineData(2, 5, 2, "Bilet miesięczny")]
        [InlineData(2, 11, 3, "Fast food")]
        public void CellReadingStringDataTest(int worksheetIndex, int rowIndex, int columnIndex, string expectedCellValue)
        {
            IWorksheetReader worksheetReader = new WorksheetReader(filepath);

            var cellCoordinates = new ExcelCellCoordinates(worksheetIndex, rowIndex, columnIndex);
            var cellValue = worksheetReader.GetCellText(cellCoordinates);
            cellValue
                .Should()
                .Be(expectedCellValue);
        }

        [Theory]
        [ClassData(typeof(GettingDateTimeFromCell))]
        public void CellReadingDateTimeDataTest(int worksheetIndex, int rowIndex, int columnIndex, DateTime expectedCellValue)
        {
            IWorksheetReader worksheetReader = new WorksheetReader(filepath);

            var cellCoordinates = new ExcelCellCoordinates(worksheetIndex, rowIndex, columnIndex);
            var cellValue = worksheetReader.GetCellDateTimeAsUTC(cellCoordinates);
            cellValue
                .Should()
                .Be(expectedCellValue);
        }

        [Theory]
        [InlineData(0, 1, 5)]
        [InlineData(0, 11, 1)]
        [InlineData(0, 2, 2)]
        [InlineData(0, 3, 4)]
        public void TryingReadingDateTimeFromCellWithAnotherDataTypeTest(int worksheetIndex, int rowIndex, int columnIndex)
        {
            IWorksheetReader worksheetReader = new WorksheetReader(filepath);

            var cellCoordinates = new ExcelCellCoordinates(worksheetIndex, rowIndex, columnIndex);
            Action getCellDateTimeAction = () => worksheetReader.GetCellDateTimeAsUTC(cellCoordinates);

            getCellDateTimeAction
                .Should()
                .ThrowExactly<WorksheetReaderCellValueTypeException>("trying to read cell with non-dateTime format should raise an exception")
                .WithMessage($"*worksheet index: {cellCoordinates.WorksheetIndex},* " +
                    $"*row index: {cellCoordinates.RowIndex},*" +
                    $"*column index {cellCoordinates.ColumnIndex}*" +
                    $"*Expected type of cell's value: {typeof(DateTime)}*");
        }

        public class GettingCellDataFromIncorrectRowOrColumnIndexesData : IEnumerable<object[]>
        {
            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public IEnumerator<object[]> GetEnumerator()
            {
                var reader = new WorksheetReader(filepath);

                Action<ExcelCellCoordinates> getCellTextFunc = coordinates => reader.GetCellText(coordinates);
                Action<ExcelCellCoordinates> getCellDateTimeFunc = coordinates => reader.GetCellDateTimeAsUTC(coordinates);

                var getCellDataFuncs = new Action<ExcelCellCoordinates>[] { getCellTextFunc, getCellDateTimeFunc };

                foreach (var func in getCellDataFuncs)
                {
                    yield return new object[] { 0, func };
                    yield return new object[] { -1, func };
                    yield return new object[] { -10, func };
                }
            }
        }

        public class GettingCellDataFromNonExistingWorksheetData : IEnumerable<object[]>
        {
            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public IEnumerator<object[]> GetEnumerator()
            {
                var reader = new WorksheetReader(filepath);

                Action<ExcelCellCoordinates> getCellTextFunc = coordinates => reader.GetCellText(coordinates);
                Action<ExcelCellCoordinates> getCellDateTimeFunc = coordinates => reader.GetCellDateTimeAsUTC(coordinates);

                var getCellDataFuncs = new Action<ExcelCellCoordinates>[] { getCellTextFunc, getCellDateTimeFunc };

                foreach (var func in getCellDataFuncs)
                {
                    yield return new object[] { -10, func };
                    yield return new object[] { -1, func };
                    yield return new object[] { 3, func };
                    yield return new object[] { 4, func };
                    yield return new object[] { 10, func };
                }
            }
        }

        public class GettingDateTimeFromCell : IEnumerable<object[]>
        {
            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public IEnumerator<object[]> GetEnumerator()
            {
                var reader = new WorksheetReader(filepath);

                yield return new object[] { 0, 19, 1, new DateTime(year: 1999, month: 6, day: 12, hour: 13, minute: 46, second: 27) };
                yield return new object[] { 0, 17, 1, new DateTime(year: 2012, month: 3, day: 8, hour: 6, minute: 15, second: 0) };
                yield return new object[] { 0, 4, 1, new DateTime(year: 2017, month: 10, day: 25) };
            }
        }

        [Theory]
        [InlineData(0, 4, 1, true)]
        [InlineData(0, 11, 1, false)]
        [InlineData(0, 13, 1, true)]
        [InlineData(0, 17, 1, true)]
        [InlineData(0, 19, 1, true)]
        [InlineData(0, 5, 1, false)]
        [InlineData(0, 1, 2, false)]
        [InlineData(0, 4, 2, false)]
        [InlineData(0, 2, 4, false)]
        public void IsCellDateTimeTest(int worksheetIndex, int rowIndex, int columnIndex, bool expectedResult)
        {
            var reader = new WorksheetReader(filepath);
            var result = reader.CellAtCoordinatesContainsDateTime(
                 new ExcelCellCoordinates(worksheetIndex, rowIndex, columnIndex));

            result
            .Should()
            .Be(expectedResult);
        }
    }
}