namespace EonUk.Retail.Ice.NonSupplyCustomerServicingModule.Helpers.Office.Excel
{
    using System;
    using System.Globalization;
    using System.Linq;
    using System.Text.RegularExpressions;

    public abstract class ExcelAddress
    {
        private static readonly Regex ColumnNameRegex = new Regex("[A-Za-z]+");
        private static readonly Regex RowIndexRegex = new Regex(@"\d+");
        private static readonly Regex AlphaNumericRegex = new Regex("^[A-Z]+$");

        public abstract ExcelAddress MoveLeft();
        public abstract ExcelAddress MoveRight();
        public abstract ExcelAddress MoveUp();
        public abstract ExcelAddress MoveDown();
        public abstract ExcelAddress MoveToRow(uint rowIndex);

        public abstract string ReferenceString { get; }
        public abstract string CellReferenceString { get; }

        public uint ColumnIndex { get; private set; }
        public uint RowIndex { get; private set; }

        public sealed class Cell : ExcelAddress
        {
            public string Column { get; private set; }

            public override string ReferenceString { get { return this.Column + this.RowIndex.ToString(CultureInfo.InvariantCulture); } }

            public override string CellReferenceString
            {
                get { return this.ReferenceString; }
            }

            public Cell(string col, uint rowIndex)
            {
                this.Column = col;
                this.RowIndex = rowIndex;
                this.ColumnIndex = GetColumnIndex(col);
            }

            public Cell(uint col, uint rowIndex)
            {
                this.Column = GetColumnName(col);
                this.RowIndex = rowIndex;
                this.ColumnIndex = col;
            }

            public override ExcelAddress MoveToRow(uint rowIndex)
            {
                return new Cell(this.Column, rowIndex);

            }

            public override ExcelAddress MoveLeft()
            {
                var colIndex = this.ColumnIndex + 1;
                return new Cell(GetColumnName(colIndex), this.RowIndex);
            }

            public override ExcelAddress MoveRight()
            {
                if (this.ColumnIndex >= 1)
                {
                    var colIndex = this.ColumnIndex - 1;
                    return new Cell(GetColumnName(colIndex), this.RowIndex);
                }

                return this;
            }

            public override ExcelAddress MoveUp()
            {
                if (this.RowIndex >= 1)
                {
                    return new Cell(this.Column, Math.Max(0, this.RowIndex - 1));
                }

                return this;
            }

            public override ExcelAddress MoveDown()
            {
                return new Cell(this.Column, this.RowIndex + 1);
            }
        }

        public class CellRange : ExcelAddress
        {
            public Cell StartCell { get; private set; }
            public Cell EndCell { get; private set; }

            public override string ReferenceString
            {
                get
                {
                    return String.Join(":", new []{this.StartCell, this.EndCell}.Select(r => r.ReferenceString).ToArray());
                }
            }

            public override string CellReferenceString
            {
                get { return this.StartCell.ReferenceString; }
            }

            public CellRange(Cell start, Cell end)
            {
                this.StartCell = start;
                this.EndCell = end;
                this.ColumnIndex = this.StartCell.ColumnIndex;
                this.RowIndex = this.StartCell.RowIndex;
            }

            public override ExcelAddress MoveToRow(uint rowIndex)
            {
                var diff = (this.StartCell.RowIndex - rowIndex);
                return new CellRange((Cell)this.StartCell.MoveToRow(rowIndex), (Cell)this.EndCell.MoveToRow(this.EndCell.RowIndex - diff));

            }

            public override ExcelAddress MoveLeft()
            {
                return new CellRange((Cell)this.StartCell.MoveLeft(), (Cell)this.EndCell.MoveLeft());
            }

            public override ExcelAddress MoveRight()
            {
                return new CellRange((Cell)this.StartCell.MoveRight(), (Cell)this.EndCell.MoveRight());
            }

            public override ExcelAddress MoveUp()
            {
                return new CellRange((Cell)this.StartCell.MoveUp(), (Cell)this.EndCell.MoveUp());
            }

            public override ExcelAddress MoveDown()
            {
                return new CellRange((Cell)this.StartCell.MoveDown(), (Cell)this.EndCell.MoveDown());
            }
        }

        public sealed class NamedRange : CellRange
        {
            public string Name { get; set; }
            public string SheetName { get; private set; }

            public override string ReferenceString
            {
                get
                {
                    return this.SheetName + "!" + base.ReferenceString;
                }
            }

            public NamedRange(string name, string sheetName, Cell start, Cell end)
                : base(start, end)
            {
                this.Name = name;
                this.SheetName = sheetName;
            }

            public override ExcelAddress MoveLeft()
            {
                return new NamedRange(this.Name, this.SheetName, (Cell)this.StartCell.MoveLeft(), (Cell)this.EndCell.MoveLeft());
            }

            public override ExcelAddress MoveRight()
            {
                return new NamedRange(this.Name, this.SheetName, (Cell)this.StartCell.MoveRight(), (Cell)this.EndCell.MoveRight());
            }

            public override ExcelAddress MoveUp()
            {
                return new NamedRange(this.Name, this.SheetName, (Cell)this.StartCell.MoveUp(), (Cell)this.EndCell.MoveUp());
            }

            public override ExcelAddress MoveDown()
            {
                return new NamedRange(this.Name, this.SheetName, (Cell)this.StartCell.MoveDown(), (Cell)this.EndCell.MoveDown());
            }
        }

        public T Match<T>(Func<Cell, T> cellAddr, Func<CellRange, T> cellRAddr)
        {
            var a = this as Cell;
            if (a != null)
            {
                return cellAddr(a);
            }

            var b = this as CellRange;
            if (b != null)
            {
                return cellRAddr(b);
            }

            throw new InvalidOperationException("Incomplete match");
        }

        public static NamedRange ParseNamedRange(string name, string definition)
        {
            var components = definition.Split('!');

            var sheetName = components[0].Trim('\'');
            //Assumption: None of my defined names are relative defined names (i.e. A1)
            string range = components[1];
            string[] rangeArray = range.Split('$');
            var start = new Cell(rangeArray[1], uint.Parse(rangeArray[2].TrimEnd(':')));
            var end = new Cell(rangeArray[1], uint.Parse(rangeArray[2].TrimEnd(':')));
            return new NamedRange(name, sheetName, start, end);
        }

        public static Cell ParseCellAddress(string defintion)
        {
            var colName = GetColumnName(defintion);
            return new Cell(colName, GetRowIndex(defintion));
        }

        public static CellRange ParseCellRange(string defintion)
        {
               var components = defintion.Split(':');

               if (components.Length == 0)
               {
                   throw new ArgumentException("Expected reference of form A1:B2");
               }

               var addrs = components.Select(c =>
               {
                   var colHeader = GetColumnName(c);
                   return new Cell(colHeader, GetRowIndex(c));
               }).ToArray();
               return new CellRange(addrs[0], addrs[1]);
        }

        public static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.

            Match match = ColumnNameRegex.Match(cellName);

            return match.Value;
        }

        public static string GetColumnName(uint index)
        {
            var intFirstLetter = ((index) / 676) + 64;
            var intSecondLetter = ((index % 676) / 26) + 64;
            var intThirdLetter = (index % 26) + 65;
	
            var firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
            var secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : ' ';
            var thirdLetter = (char)intThirdLetter;
	
            return String.Concat(firstLetter, secondLetter, thirdLetter).Trim();
        }

        public static uint GetColumnIndex(string column)
        {

            if (!AlphaNumericRegex.IsMatch(column)) throw new ArgumentException();

            char[] colLetters = column.ToCharArray();
            Array.Reverse(colLetters);

            uint convertedValue = 0;
            for (uint i = 0; i < colLetters.Length; i++)
            {
                char letter = colLetters[i];
                // ASCII 'A' = 65
                var current = (uint)(i == 0 ? letter - 65 : letter - 64);
                convertedValue += current * (uint)Math.Pow(26, i);
            }

            return convertedValue;
        }

        public static uint GetRowIndex(string cellReference)
        {
            Match match = RowIndexRegex.Match(cellReference);

            return UInt32.Parse(match.Value);
        }


    }
}