using System.Collections.Generic;

namespace DataTransferer.Lib.FileReader.DataParser
{
    public interface IDataParser<T>
    {
        IEnumerable<T> ParseData();
    }
}