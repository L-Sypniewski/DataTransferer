using System.Collections.Generic;
using SpendingsDataTransferer.Lib.ApplicationModel.Excel;

namespace SpendingsDataTransferer.Lib.FileReader.Excel.WorksheetReader
{
    public interface IWorksheetReader
    {
        IEnumerable<Worksheet> Worksheets { get; }

        T GetCellValue<T>(int worksheetIndex, int row, int column);
    }
}