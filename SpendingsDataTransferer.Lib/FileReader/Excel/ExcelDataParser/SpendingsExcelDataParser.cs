using SpendingsDataTransferer.Lib.ApplicationModel;
using SpendingsDataTransferer.Lib.FileReader.Excel.WorksheetReader;

namespace SpendingsDataTransferer.Lib.FileReader.Excel.ExcelDataParser
{
    public class SpendingsExcelDataParser : IExcelDataParser<Spending>
    {
        private readonly string filename;
        private readonly IWorksheetReader worksheetReader;

        public SpendingsExcelDataParser(string filename, IWorksheetReader worksheetReader) {
            this.filename = filename;
            this.worksheetReader = worksheetReader;
        }
        public Spending ParseData()
        {
            throw new System.NotImplementedException();
        }
    }
}