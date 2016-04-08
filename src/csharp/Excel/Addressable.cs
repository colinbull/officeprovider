namespace Office.Excel
{
    using System;

    using DocumentFormat.OpenXml.Spreadsheet;

    public class Addressable<T, TAddress>
        where TAddress : ExcelAddress
    {
        public T Item { get; private set; }
        public TAddress Address { get; private set; }

        public Addressable(T item, Func<T, TAddress> parseFunction)
            :this(item, parseFunction(item))
        {
        }

        private Addressable(T item, TAddress address)
        {
            this.Item = item;
            this.Address = address;
        }

        public static Addressable<Cell, ExcelAddress.Cell> Lift(Cell item)
        {
            return new Addressable<Cell, ExcelAddress.Cell>(item, c => ExcelAddress.ParseCellAddress(c.CellReference));
        }

        public static Addressable<MergeCell, ExcelAddress.CellRange> Lift(MergeCell item)
        {
            return new Addressable<MergeCell, ExcelAddress.CellRange>(item, c => ExcelAddress.ParseCellRange(c.Reference));
        }

        public static Addressable<DefinedName, ExcelAddress.NamedRange> Lift(DefinedName item)
        {
            return new Addressable<DefinedName, ExcelAddress.NamedRange>(item, c => ExcelAddress.ParseNamedRange(item.Name.Value, c.InnerText));
        }

        public Addressable<T, TAddress> MoveLeft()
        {
            return new Addressable<T, TAddress>(this.Item, (TAddress)this.Address.MoveLeft());
        }

        public Addressable<T, TAddress> MoveRight()
        {
            return new Addressable<T, TAddress>(this.Item, (TAddress)this.Address.MoveRight());
        }

        public Addressable<T, TAddress> MoveUp()
        {
            return new Addressable<T, TAddress>(this.Item, (TAddress)this.Address.MoveUp());
        }

        public Addressable<T, TAddress> MoveDown()
        {
            return new Addressable<T, TAddress>(this.Item, (TAddress)this.Address.MoveDown());
        }

    }
}