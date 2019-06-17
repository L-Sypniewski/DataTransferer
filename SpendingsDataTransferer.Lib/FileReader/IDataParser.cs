using System.Collections.Generic;

namespace SpendingsDataTransferer.Lib.FileReader.DataParser
{
    public interface IDataParser<T>
    {
         IEnumerable<T> ParseData();
    }
}