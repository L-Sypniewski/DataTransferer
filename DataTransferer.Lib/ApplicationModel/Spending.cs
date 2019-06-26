using System;

namespace DataTransferer.Lib.ApplicationModel
{
    public struct Spending: IComparable<Spending>, IEquatable<Spending>
    {
        public readonly DateTime Date;
        public readonly string Name;
        public readonly string Category;
        public readonly Decimal Amount;
        public readonly Currency Currency;

        public Spending(DateTime date, string name, string category, decimal amount, Currency currency)
        {
            Date = date;
            Name = name;
            Category = category;
            Amount = amount;
            Currency = currency;
        }

        public int CompareTo(Spending other)
        {
            throw new NotImplementedException();
        }

        public bool Equals(Spending other)
        {
            throw new NotImplementedException();
        }
    }
}