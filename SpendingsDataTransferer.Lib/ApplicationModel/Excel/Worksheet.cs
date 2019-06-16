namespace SpendingsDataTransferer.Lib.ApplicationModel.Excel
{
    public struct Worksheet
    {
        public string Name { get; }

        public Worksheet(string name)
        {
            this.Name = name;
        }
    }
}