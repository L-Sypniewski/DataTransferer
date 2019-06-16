namespace SpendingsDataTransferer.Lib.FileReader.Excel.ExcelDataParser
{
    public interface IExcelDataParser<T>
    {
         T ParseData();
    }
}