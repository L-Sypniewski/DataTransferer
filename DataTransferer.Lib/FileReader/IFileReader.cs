using System.Collections.Generic;

namespace DataTransferer.Lib.FileReader
{
    public interface IFileReader<T>
    {
        IEnumerable<T> GetData();
    }
}