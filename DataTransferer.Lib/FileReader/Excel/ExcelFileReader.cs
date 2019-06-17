using System.Collections.Generic;
using DataTransferer.Lib.FileReader.DataParser;

namespace DataTransferer.Lib.FileReader.Excel
{
    public class ExcelFileReader<T> : IFileReader<T>
    {
        private readonly string filename;
        private readonly IDataParser<T> dataParser;

        public ExcelFileReader(string filename, IDataParser<T> dataParser)
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