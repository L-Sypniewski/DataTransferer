using System;
using System.Collections.Generic;
using SpendingsDataTransferer.Lib.ApplicationModel.Excel;

namespace SpendingsDataTransferer.Lib.FileReader.Excel.WorksheetReader
{
    public interface IWorksheetReader
    {
        IEnumerable<Worksheet> Worksheets { get; }

        string GetCellText(ExcelCellCoordinates cellCoordinates);
        DateTime GetCellDateTime(ExcelCellCoordinates cellCoordinates);
    }
}