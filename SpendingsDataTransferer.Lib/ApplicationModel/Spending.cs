using System;

namespace SpendingsDataTransferer.Lib.ApplicationModel
{
    public struct Spending
    {
        DateTime Date;
        string Name;
        string Category;
        Decimal Amount;
        string Currency;
    }
}