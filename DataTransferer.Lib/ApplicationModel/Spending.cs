using System;

namespace DataTransferer.Lib.ApplicationModel
{
    public struct Spending
    {
        DateTime Date;
        string Name;
        string Category;
        Decimal Amount;
        Currency Currency;
    }
}