using System;
using System.Collections.Generic;
using DataTransferer.Lib.ApplicationModel.Excel;

namespace DataTransferer.Lib.FileReader.Excel.WorksheetReader
{
    public interface IWorksheetReader
    {
        IEnumerable<Worksheet> Worksheets { get; }

        string GetCellText(ExcelCellCoordinates cellCoordinates);
        DateTime GetCellDateTime(ExcelCellCoordinates cellCoordinates);
    }
}