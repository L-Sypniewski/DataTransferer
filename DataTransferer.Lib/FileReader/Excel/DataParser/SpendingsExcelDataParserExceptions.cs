using System;
using System.Runtime.Serialization;
using DataTransferer.Lib.ApplicationModel.Excel;

namespace DataTransferer.Lib.FileReader.Excel.WorksheetReader
{
    [Serializable]
    public class MissingSpendingsItemException : Exception
    {
        private static string ExceptionSpecificMessage(string missingItemName, ExcelCellCoordinates cellCoordinates) =>
            $"There is no {missingItemName} at index:\n" +
            $"worksheet: {cellCoordinates.WorksheetIndex}\nrow: {cellCoordinates.RowIndex}\ncolumn: {cellCoordinates.ColumnIndex}";

        public MissingSpendingsItemException() { }
        public MissingSpendingsItemException(string missingItemName, ExcelCellCoordinates cellCoordinates)
            : base(ExceptionSpecificMessage(missingItemName, cellCoordinates)) { }
        public MissingSpendingsItemException(string missingItemName, ExcelCellCoordinates cellCoordinates, Exception inner)
            : base(ExceptionSpecificMessage(missingItemName, cellCoordinates), inner) { }
        protected MissingSpendingsItemException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}