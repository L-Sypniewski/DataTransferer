using System;
using System.Runtime.Serialization;
using DataTransferer.Lib.ApplicationModel.Excel;

namespace DataTransferer.Lib.FileReader.Excel.WorksheetReader
{
    [Serializable]
    public class WorksheetReaderNonExistingWorksheetException : Exception
    {
        private static string exceptionSpecificMessage(string index) => $"There is no worksheet at index: {index}";
        public WorksheetReaderNonExistingWorksheetException() { }
        public WorksheetReaderNonExistingWorksheetException(string index) : base(exceptionSpecificMessage(index)) { }
        public WorksheetReaderNonExistingWorksheetException(
            string index, Exception inner) : base(exceptionSpecificMessage(index), inner) { }
        protected WorksheetReaderNonExistingWorksheetException(
            SerializationInfo info,
            StreamingContext context) : base(info, context) { }
    }

    public class WorksheetReaderColumnLesserThanOneException : Exception
    {
        private static string exceptionSpecificMessage(string index) => $"WorksheetReader tried to read column at index: {index}";
        public WorksheetReaderColumnLesserThanOneException() { }
        public WorksheetReaderColumnLesserThanOneException(string index) : base(exceptionSpecificMessage(index)) { }
        public WorksheetReaderColumnLesserThanOneException(
            string index, Exception inner) : base(exceptionSpecificMessage(index), inner) { }
        protected WorksheetReaderColumnLesserThanOneException(
            SerializationInfo info,
            StreamingContext context) : base(info, context) { }
    }

    public class WorksheetReaderRowLesserThanOneException : Exception
    {
        private static string exceptionSpecificMessage(string index) => $"WorksheetReader tried to read row at index: {index}";
        public WorksheetReaderRowLesserThanOneException() { }
        public WorksheetReaderRowLesserThanOneException(string index) : base(exceptionSpecificMessage(index)) { }
        public WorksheetReaderRowLesserThanOneException(
            string index, Exception inner) : base(exceptionSpecificMessage(index), inner) { }
        protected WorksheetReaderRowLesserThanOneException(
            SerializationInfo info,
            StreamingContext context) : base(info, context) { }
    }

    public class WorksheetReaderCellValueTypeException : Exception
    {
        private static string exceptionSpecificMessage(ExcelCellCoordinates cellCoordinates, Type expectedType) =>
            $"WorksheetReader tried to read row form worksheet index: {cellCoordinates.WorksheetIndex}, " +
            $"row index: {cellCoordinates.RowIndex}, column index {cellCoordinates.ColumnIndex}\n" +
            $"Expected type of cell's value: {expectedType}";
        public WorksheetReaderCellValueTypeException() { }
        public WorksheetReaderCellValueTypeException(ExcelCellCoordinates cellCoordinates, Type expectedType) :
            base(exceptionSpecificMessage(cellCoordinates, expectedType))
        { }
        public WorksheetReaderCellValueTypeException(
                ExcelCellCoordinates cellCoordinates, Type expectedType, Exception inner) :
            base(exceptionSpecificMessage(cellCoordinates, expectedType), inner)
        { }
        protected WorksheetReaderCellValueTypeException(
            SerializationInfo info,
            StreamingContext context) : base(info, context) { }
    }

}