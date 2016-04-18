namespace Office.Excel
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text.RegularExpressions;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    public class ExcelWorksheet
    {
        private readonly WorksheetPart worksheetPart;
        private readonly Dictionary<uint, Dictionary<uint, Cell>> cells = new Dictionary<uint, Dictionary<uint, Cell>>();
        public ExcelWorkbook Workbook { get; private set; }

        public string Name { get; private set; }

        public WorksheetPart WorksheetPart
        {
            get { return this.worksheetPart; }
        }

        public ExcelWorksheet(ExcelWorkbook workbook, string name, WorksheetPart part)
        {
            this.Workbook = workbook;
            this.Name = name;
            this.worksheetPart = part;
            this.BuildCellIndex();
        }

        public Cell GetCell(ExcelAddress.Cell address)
        {
            Dictionary<uint, Cell> rows;
            if (this.cells.TryGetValue(address.ColumnIndex, out rows))
            {
                Cell cell;
                if (rows.TryGetValue(address.RowIndex, out cell))
                {
                    return cell;
                }
            }

            throw new KeyNotFoundException(String.Format("No cell exists at {0}", address.CellReferenceString));
        }

        public ExcelWorksheet Clone(string newName)
        {
            return this.Workbook.CopySheet(this.Name, newName);
        }

        public void WriteValue(ExcelAddress.Cell address, object value)
        {
            var cell = this.GetCell(address);
            if (value == null)
            {
                cell.CellValue = new CellValue(null);
                return;
            }

            var typ = value.GetType();
            if (typ == typeof(string))
            {
                var v = (string)value;
                var index = this.Workbook.InsertSharedStringItem(v);
                cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                return;
            }

            if (typ == typeof(DateTime))
            {
                var dt = (DateTime)value;
                string dtValue = dt.ToOADate().ToString(CultureInfo.InvariantCulture);
                cell.CellValue = new CellValue(dtValue);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                
                return;
            }

            if (ExcelWorkbook.NumericTypes.Contains(typ))
            {
                cell.CellValue = new CellValue(value.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                return;
            }

            throw new InvalidOperationException(string.Format("the type {0} is not currently supported if it is an array values should be boxed to type object[][]", typ.FullName));
        }

        public void WriteArray(Addressable<DefinedName, ExcelAddress.NamedRange> address, string[] value)
        {
            var startCell = address.Address;
            var currentCell = startCell;
            var sheet = this;

            foreach (var x in value)
            {
                sheet.WriteValue(currentCell.StartCell, x);
                currentCell = (ExcelAddress.NamedRange)currentCell.MoveLeft();
            }
        }

        public void WriteArray2D(Addressable<DefinedName, ExcelAddress.NamedRange> address, string[][] value)
        {
            var startCell = address.Address;
            var currentCell = startCell;
            var sheet = this;

            foreach (var ys in value)
            {
                foreach (var x in ys)
                {
                    sheet.WriteValue(currentCell.StartCell, x);
                    currentCell = (ExcelAddress.NamedRange)currentCell.MoveLeft();
                }

                currentCell = (ExcelAddress.NamedRange)currentCell.MoveDown();
            }
        }

        public void WritePagedArray(Addressable<DefinedName, ExcelAddress.NamedRange> address, ExcelPagedArray value)
        {
            var startCell = address.Address;
            var startRowIndex = startCell.RowIndex;
            var currentCell = startCell;
            var sheet = this;


            sheet.CopyRow(startRowIndex, value.PageSize);

            var rowsWrittenToThisSheet = 0;
            var sheetNumber = 0;

            for (var i = 0; i < value.Length; i++)
            {
                foreach (var x in value[i])
                {
                    sheet.WriteValue(currentCell.StartCell, x);
                    currentCell = (ExcelAddress.NamedRange)currentCell.MoveLeft();
                }

                rowsWrittenToThisSheet++;
                if (rowsWrittenToThisSheet < value.PageSize)
                {
                    startCell = (ExcelAddress.NamedRange)startCell.MoveDown();
                    currentCell = startCell;
                }
                else
                {
                    if ((i + 1) < value.Length)
                    {
                        sheetNumber++;

                        sheet = sheet.Clone(sheet.Name + "_" + sheetNumber);

                        sheet.CopyRow(address.Address.RowIndex, (int)Math.Min(value.PageSize, value.Length - i - 1));
                        
                        startCell = address.Address;
                        currentCell = startCell;
                        rowsWrittenToThisSheet = 0;
                    }
                }
            }
        }

        public void CopyRow(uint rowIndex, int copies)
        {
            for (var r = 1; r < copies; r++)
            {
                this.InsertCopyOfRow(rowIndex, (uint)(rowIndex + r), false);
            }

            this.BuildCellIndex();
        }

        public void InsertRow(uint rowIndex, Row insertRow, bool rebuildIndex)
        {
            this.InsertRowInternal(rowIndex, insertRow, false, rebuildIndex);
        }

        public void AppendRow(uint rowIndex, Row insertRow, bool rebuildIndex)
        {
            this.InsertRowInternal(rowIndex, insertRow, true, rebuildIndex);
        }

        public void InsertCopyOfRow(uint rowIndex, uint targetRowIndex, bool rebuildIndex)
        {
            var sheetData = this.WorksheetPart.Worksheet.GetFirstChild<SheetData>();

            var rows = sheetData.Elements<Row>();
            var lastRow = rows.Last().RowIndex;
            var currRow = rows.FirstOrDefault(r => r.RowIndex == rowIndex);

            if (currRow != null)
            {
                var newRow = (Row)currRow.CloneNode(true);

                if (targetRowIndex >= lastRow)
                {
                    this.AppendRow(targetRowIndex, newRow, rebuildIndex);
                }
                else
                {
                    this.InsertRow(targetRowIndex, newRow, rebuildIndex);
                }

                var mergeCells = this.WorksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
                if (mergeCells != null)
                {
                    var rowMergedCells = mergeCells.Elements<MergeCell>()
                        .Select(r => Addressable<MergeCell, ExcelAddress.CellRange>.Lift(r))
                        .Where(r => r.Address.Match(c => c.RowIndex == rowIndex, c => c.StartCell.RowIndex == rowIndex && c.EndCell.RowIndex == rowIndex))
                        .Select(r =>
                            {
                                var cell = (MergeCell)r.Item.CloneNode(true);
                                cell.Reference = r.Address.MoveToRow(targetRowIndex).ReferenceString;
                                return cell;
                            });

                    mergeCells.Append(rowMergedCells.Cast<OpenXmlElement>());
                }
            }
        }

        public void BuildCellIndex()
        {
            this.cells.Clear();

            foreach (var cell in this.WorksheetPart.Worksheet.Descendants<Cell>())
            {
                this.EnsureCellIsInLookup(cell);
            }

        }

        private void EnsureCellIsInLookup(Cell cell)
        {
            var address = ExcelAddress.ParseCellAddress(cell.CellReference);
            if (!this.cells.ContainsKey(address.ColumnIndex))
            {
                this.cells[address.ColumnIndex] = new Dictionary<uint, Cell>();
            }

            if (!this.cells[address.ColumnIndex].ContainsKey(address.RowIndex))
            {
                this.cells[address.ColumnIndex][address.RowIndex] = cell;
            }
        }

        private Row InsertRowInternal(uint rowIndex, Row insertRow, bool isNewLastRow, bool rebuildIndex)
        {
            Worksheet worksheet = this.WorksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            Row retRow = !isNewLastRow ? sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex) : null;

            // If the worksheet does not contain a row with the specified row index, insert one.
            if (retRow != null)
            {
                // if retRow is not null and we are inserting a new row, then move all existing rows down.
                if (insertRow != null)
                {
                    this.UpdateRowIndexes(rowIndex, false);
                    this.UpdateMergedCellReferences(rowIndex, false);
                    this.UpdateHyperlinkReferences(rowIndex, false);

                    // actually insert the new row into the sheet
                    retRow = sheetData.InsertBefore(insertRow, retRow);  // at this point, retRow still points to the row that had the insert rowIndex

                    string curIndex = retRow.RowIndex.ToString();
                    string newIndex = rowIndex.ToString();

                    foreach (Cell cell in retRow.Elements<Cell>())
                    {
                        // Update the references for the rows cells.
                        cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curIndex, newIndex));
                    }

                    // Update the row index.
                    retRow.RowIndex = rowIndex;
                }
            }
            else
            {
                // Row doesn't exist yet, shifting not needed.
                // Rows must be in sequential order according to RowIndex. Determine where to insert the new row.
                Row refRow = !isNewLastRow ? sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex > rowIndex) : null;

                // use the insert row if it exists
                retRow = insertRow ?? new Row() { RowIndex = rowIndex };

                IEnumerable<Cell> cellsInRow = retRow.Elements<Cell>();

                if (cellsInRow.Any())
                {
                    string curIndex = retRow.RowIndex.ToString();
                    string newIndex = rowIndex.ToString();

                    foreach (Cell cell in cellsInRow)
                    {
                        // Update the references for the rows cells.
                        cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curIndex, newIndex));
                    }

                    // Update the row index.
                    retRow.RowIndex = rowIndex;
                }

                sheetData.InsertBefore(retRow, refRow);
            }

            if (rebuildIndex)
            {
                this.BuildCellIndex();
            }

            return retRow;
        }

        private void UpdateRowIndexes(uint rowIndex, bool isDeletedRow)
        {
            // Get all the rows in the worksheet with equal or higher row index values than the one being inserted/deleted for reindexing.
            IEnumerable<Row> rows = this.WorksheetPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= rowIndex);

            foreach (Row row in rows)
            {
                uint newIndex = (isDeletedRow ? row.RowIndex - 1 : row.RowIndex + 1);
                string curRowIndex = row.RowIndex.ToString();
                string newRowIndex = newIndex.ToString();

                foreach (Cell cell in row.Elements<Cell>())
                {
                    // Update the references for the rows cells.
                    cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curRowIndex, newRowIndex));
                }

                // Update the row index.
                row.RowIndex = newIndex;
            }
        }

        private void UpdateMergedCellReferences(uint rowIndex, bool isDeletedRow)
        {
            if (this.WorksheetPart.Worksheet.Elements<MergeCells>().Any())
            {
                MergeCells mergeCells = this.WorksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();

                if (mergeCells != null)
                {
                    var mergeCellsList =
                        mergeCells.Elements<MergeCell>()
                            .Select(r => Addressable<MergeCell, ExcelAddress.CellRange>.Lift(r))
                            .Where(r => r.Address.Match(c => c.RowIndex >= rowIndex,
                                c => c.StartCell.RowIndex >= rowIndex || c.EndCell.RowIndex >= rowIndex)
                            );

                    if (isDeletedRow)
                    {
                        var mergeCellsToDelete =
                            mergeCellsList.Where(r =>
                                                 r.Address.Match(c => c.RowIndex == rowIndex,
                                                     c => c.StartCell.RowIndex == rowIndex || c.EndCell.RowIndex == rowIndex));

                        foreach (var cellToDelete in mergeCellsToDelete)
                        {
                            cellToDelete.Item.Remove();
                        }

                        // Update the list to contain all merged cells greater than the deleted row index
                        mergeCellsList =
                            mergeCells
                                .Elements<MergeCell>()
                                .Select(r => Addressable<MergeCell, ExcelAddress.CellRange>.Lift(r))
                                .Where(r =>
                                       r.Address.Match(c => c.RowIndex > rowIndex,
                                           c => c.StartCell.RowIndex > rowIndex || c.EndCell.RowIndex > rowIndex))
                                .ToList();
                    }

                    // Either increment or decrement the row index on the merged cell reference
                    foreach (var mergeCell in mergeCellsList.ToArray())
                    {
                        var addr = isDeletedRow ? mergeCell.Address.MoveUp() : mergeCell.Address.MoveDown();
                        mergeCell.Item.Reference = addr.ReferenceString;
                    }
                }
            }
        }

        /// <summary>
        /// Updates all hyperlinks in the worksheet when a row is inserted or deleted.
        /// </summary>
        /// <param name="worksheetPart">Worksheet Part</param>
        /// <param name="rowIndex">Row Index being inserted or deleted</param>
        /// <param name="isDeletedRow">True if row was deleted, otherwise false</param>
        private void UpdateHyperlinkReferences(uint rowIndex, bool isDeletedRow)
        {
            Hyperlinks hyperlinks = this.WorksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();

            if (hyperlinks != null)
            {
                Match hyperlinkRowIndexMatch;
                uint hyperlinkRowIndex;

                foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>())
                {
                    hyperlinkRowIndexMatch = Regex.Match(hyperlink.Reference.Value, "[0-9]+");
                    if (hyperlinkRowIndexMatch.Success && UInt32.TryParse(hyperlinkRowIndexMatch.Value, out hyperlinkRowIndex) && hyperlinkRowIndex >= rowIndex)
                    {
                        // if being deleted, hyperlink needs to be removed or moved up
                        if (isDeletedRow)
                        {
                            // if hyperlink is on the row being removed, remove it
                            if (hyperlinkRowIndex == rowIndex)
                            {
                                hyperlink.Remove();
                            }
                            // else hyperlink needs to be moved up a row
                            else
                            {
                                hyperlink.Reference.Value = hyperlink.Reference.Value.Replace(hyperlinkRowIndexMatch.Value, (hyperlinkRowIndex - 1).ToString());

                            }
                        }
                        // else row is being inserted, move hyperlink down
                        else
                        {
                            hyperlink.Reference.Value = hyperlink.Reference.Value.Replace(hyperlinkRowIndexMatch.Value, (hyperlinkRowIndex + 1).ToString());
                        }
                    }
                }

                // Remove the hyperlinks collection if none remain
                if (hyperlinks.Elements<Hyperlink>().Count() == 0)
                {
                    hyperlinks.Remove();
                }
            }
        }

    }
}