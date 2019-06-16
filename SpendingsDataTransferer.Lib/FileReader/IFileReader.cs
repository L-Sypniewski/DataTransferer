using System.Collections.Generic;

namespace SpendingsDataTransferer.Lib.FileReader
{
    public interface IFileReader<T>
    {
        IEnumerable<T>  GetData();
    }
}