namespace DataTransferer.Lib.ApplicationModel.Excel
{
    public struct ExcelCellCoordinates
    {
        public readonly int WorksheetIndex;
        public readonly int RowIndex;
        public readonly int ColumnIndex;

        public ExcelCellCoordinates(int worksheetIndex, int rowIndex, int columnIndex)
        {
            WorksheetIndex = worksheetIndex;
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }
    }
}