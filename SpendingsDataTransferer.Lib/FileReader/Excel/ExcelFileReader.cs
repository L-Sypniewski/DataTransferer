using System.Collections.Generic;
using SpendingsDataTransferer.Lib.FileReader.Excel.ExcelDataParser;

namespace SpendingsDataTransferer.Lib.FileReader.Excel
{
    public class ExcelFileReader<T> : IFileReader<T>
    {
        private readonly string filename;
        private readonly IExcelDataParser<T> dataParser;

        public ExcelFileReader(string filename, IExcelDataParser<T> dataParser)
        {
            this.filename = filename;
            this.dataParser = dataParser;
        }

        public IEnumerable<T> GetData()
        {
            throw new System.NotImplementedException();
        }
    }
}