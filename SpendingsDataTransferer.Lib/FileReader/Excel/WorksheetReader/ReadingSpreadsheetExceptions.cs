using System;
using System.Runtime.Serialization;

namespace SpendingsDataTransferer.Lib.FileReader.Excel.WorksheetReader
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

}