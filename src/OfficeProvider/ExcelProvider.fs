namespace OfficeProvider

open System
open System.IO
open System.Linq
open System.Text.RegularExpressions
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet

type ExcelCell = {
    RowIndex : uint32
    ColumnIndex : uint32
    Column : string
}
with 
    static member Empty = { RowIndex = 0u; ColumnIndex = 0u; Column = ""; }

type ExcelAddress = 
    | Cell of sheet:string * cell:ExcelCell
    | Range of sheet:string * startCell:ExcelCell * endCell:ExcelCell
    with 
        member x.Sheet 
            with get() =
                match x with
                | Cell(sheet = s) -> s
                | Range(sheet = s) -> s
        member x.Indexes
            with get() = 
                match x with
                | Cell(cell = c) -> c.ColumnIndex, c.RowIndex
                | Range(startCell = c) -> c.ColumnIndex, c.RowIndex

type ExcelField = {
    Name : string
    Type : Type
    Sheet : string
    RowIndex : uint32
    ColumnIndex : uint32
}

module Excel = 
    
    let ColumnNameRegex = new Regex("[A-Za-z]+");
    let RowIndexRegex = new Regex(@"\d+");
    let AlphaNumericRegex = new Regex("^[A-Z]+$");
    
    let (|EOF|_|) (c : char) =
        let value = c |> int
        if (value = -1 || value = 65535) then Some c else None
    let (|Letter|_|) c = if Char.IsLetter(c) then Some c else None
    let (|Digit|_|) c = if Char.IsDigit(c) then Some c else None

    type AddressToken = 
        | Sheet of string
        | Column of string
        | Row of string
    
    let columnIndex (col:string) = 
        Array.rev (col |> Seq.toArray)
        |> Array.mapi (fun i letter -> (if i = 0 then (letter |> uint32) - 65u else (letter |> uint32) - 64u) * (Math.Pow(26., i |> float) |> uint32))
        |> Array.sum

    let tokenizeCellAddress (address:string) =
        let readBuffer buffer = new String(buffer |> List.rev |> List.toArray)
        let rec sheet (buffer:_ list) state (reader:StringReader) =
            match char(reader.Read()) with
            | '!' when (not buffer.IsEmpty) -> cell [] (Sheet(readBuffer buffer) :: state) reader
            | '!' -> cell [] state reader
            | EOF _ -> state
            | a -> sheet (a :: buffer) state reader    
        and cell buffer state reader = 
            match char(reader.Read()) with
            | '$' -> column buffer state reader
            | Digit a -> row [a] state reader
            | Letter a -> column [a] state reader
            | EOF _ -> state
            | _ -> cell buffer state reader
        and column buffer state reader =
            match char(reader.Read()) with
            | Letter a -> column (a :: buffer) state reader
            | '$' -> row [] (Column(readBuffer buffer) :: state) reader
            | Digit a -> row [a] (Column(readBuffer buffer) :: state) reader
            | EOF _ -> Column(readBuffer buffer) :: state
            | _ -> column buffer state reader
        and row buffer state reader =
            match char(reader.Read()) with
            | Digit a -> row (a :: buffer) state reader
            | EOF _ -> Row(readBuffer buffer) :: state
            | ':' -> cell [] (Row(readBuffer buffer) :: state) reader
            | _ -> row buffer state reader
        sheet [] [] (new StringReader(if address.Contains("!") then address else "!" + address))
        |> List.rev

    let parseCellAddress address =
        let (sheetName, _, cells) =
            tokenizeCellAddress address 
            |> List.fold (fun (s,stack,res) x -> 
               match x with
               | Sheet(sheetName) -> (sheetName.Trim([|'\''|]), stack, res)
               | Column(col) -> (s, col, res)
               | Row(row) -> (s,"", { Column = stack; RowIndex = uint32(row); ColumnIndex = columnIndex stack } :: res ) 
            ) ("","", [])
        match cells |> List.rev with
        | [a;b] -> Range(sheetName, a, b)
        | [a] -> Cell(sheetName, a)
        | _ -> failwithf "Unable to parse cell address %s" address

type ExcelProvider(resolutionPath:string, document:string) = 
     let documentPath = 
            if String.IsNullOrWhiteSpace(resolutionPath)
            then document
            else Path.Combine(resolutionPath, document)

     let doc =
        if File.Exists(documentPath)
        then SpreadsheetDocument.Open(documentPath, true, new OpenSettings(AutoSave = true))
        else raise(FileNotFoundException("Could not find file", documentPath))

     let definedNames = 
        doc.WorkbookPart.Workbook.DefinedNames
        |> Seq.cast<DefinedName>
        |> Seq.map (fun dn -> dn.Name.Value, Excel.parseCellAddress dn.InnerText)
        |> Map.ofSeq

     let sheets = 
        doc.WorkbookPart.Workbook.Descendants<Sheet>()
        |> Seq.map (fun s -> s.Name.Value, doc.WorkbookPart.GetPartById(s.Id.Value) :?> WorksheetPart)
        |> Map.ofSeq

     let readCellValue (cell:Cell) =
        let text = cell.CellValue.InnerText
        match cell.DataType.Value with
        | CellValues.Number -> 
            Decimal.Parse text |> box
        | CellValues.SharedString ->
            let stringTable = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>() |> Seq.head
            stringTable.SharedStringTable.ElementAt(Int32.Parse(text)).InnerText |> box
        | CellValues.Boolean -> 
            Boolean.Parse text |> box
        | CellValues.Date -> 
            DateTime.FromOADate(Double.Parse(text)) |> box
        | _ -> text |> box

     interface IOfficeProvider with
        member x.GetFields() = 
            definedNames
            |> Map.toArray 
            |> Array.map (fun (n, address) -> { FieldName = n; Type = typeof<string> })

        member x.ReadField(name:string) =
            let cell =
                let address = definedNames.[name]
                sheets.[address.Sheet].Worksheet.Descendants<Cell>()
                |> Seq.find (fun (x:Cell) ->
                    let cellAddr = (Excel.parseCellAddress x.CellReference.Value) 
                    cellAddr.Indexes = address.Indexes)
            readCellValue cell

        member x.SetField(name:string, value:obj) = ()

        member x.Dispose() = 
            doc.Close()
