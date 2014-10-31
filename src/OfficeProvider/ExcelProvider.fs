namespace OfficeProvider

open System
open System.IO
open System.Linq
open System.Globalization
open System.Collections.Generic
open System.Text.RegularExpressions
open DocumentFormat.OpenXml
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

type ExcelProvider(resolutionPath:string, document:string, shadowCopy:bool) = 
    
    let NumericTypes = 
        HashSet [
             typeof<decimal>; typeof<int8>; typeof<uint8>;
             typeof<int16>; typeof<uint16>; typeof<int32>; typeof<uint32>; typeof<int64>;
             typeof<uint64>; typeof<float>; typeof<float32>
        ]
    
    let documentPath = File.getPath resolutionPath document "xlsx" shadowCopy
    let doc = SpreadsheetDocument.Open(documentPath, true)

    let definedNames = 
       doc.WorkbookPart.Workbook.DefinedNames
       |> Seq.cast<DefinedName>
       |> Seq.map (fun dn -> dn.Name.Value, Excel.parseCellAddress dn.InnerText)
       |> Map.ofSeq

    let sheets = 
       doc.WorkbookPart.Workbook.Descendants<Sheet>()
       |> Seq.choose (fun s -> 
            match doc.WorkbookPart.GetPartById(s.Id.Value) with
            | :? WorksheetPart as a -> Some(s.Name.Value, a)
            | _ -> None)
       |> Map.ofSeq

    let getStringTable() = 
        let stringTable = 
            let st = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
            if st = null
            then doc.WorkbookPart.AddNewPart<SharedStringTablePart>()
            else st
        if stringTable.SharedStringTable = null then stringTable.SharedStringTable <- new SharedStringTable()
        stringTable.SharedStringTable

    let insertSharedString str =  
        let table = getStringTable()
        
        let rec find' index found (elements : SharedStringItem list) = 
            if found then Choice1Of2 index
            else
                match elements with
                | [] -> Choice2Of2 index
                | h :: t -> find' (index + 1) (h.InnerText = str) t 

        match find' 0 false (table.Elements<SharedStringItem>() |> Seq.toList) with
        | Choice1Of2(i) -> i
        | Choice2Of2(i) ->
            table.AppendChild(new SharedStringItem([|new Text(str)|] |> Seq.cast<OpenXmlElement>)) |> ignore
            table.Save();
            i

    let readCellValue (cell:Cell) =
        if cell.CellValue <> null
        then
            let text = cell.CellValue.InnerText
            if (cell.DataType <> null) && (cell.DataType.HasValue)
            then
                match cell.DataType.Value with
                | CellValues.Number -> 
                    Decimal.Parse text |> box
                | CellValues.SharedString ->
                    getStringTable().ElementAt(Int32.Parse(text)).InnerText |> box
                | CellValues.Boolean -> 
                    Boolean.Parse text |> box
                | CellValues.Date -> 
                    DateTime.FromOADate(Double.Parse(text)) |> box
                | _ -> text |> box
            else text |> box
        else null
        
    let writeCellValue (cell:Cell) (value:obj) = 
        match value with
        | :? string as v ->
             let index = insertSharedString(v)
             cell.CellValue <- new CellValue(index.ToString(CultureInfo.InvariantCulture))
             cell.DataType <- new EnumValue<CellValues>(CellValues.SharedString)
        | :? DateTime as v ->
            let dtStr = v.ToOADate().ToString(CultureInfo.InvariantCulture)
            cell.CellValue <- new CellValue(dtStr)
            cell.DataType <- new EnumValue<CellValues>(CellValues.Date)
        | _ when NumericTypes.Contains(value.GetType()) -> 
            cell.CellValue <- new CellValue(value.ToString())
            cell.DataType <- new EnumValue<CellValues>(CellValues.Number)
        | null -> 
            cell.CellValue <- new CellValue(null)
        | _ -> failwithf "Unable to write type %A" (value.GetType().Name)
    
    let tryGetCell (a:ExcelAddress) =
        let sheet = sheets.[a.Sheet]
        sheet.Worksheet.Descendants<Cell>()
        |> Seq.tryFind (fun (x:Cell) ->
            let cellAddr = (Excel.parseCellAddress x.CellReference.Value) 
            cellAddr.Indexes = a.Indexes)

    let getCellType = function 
        | Range _ -> typeof<obj[][]>
        | Cell _ as a -> 
            match tryGetCell a with
            | Some(cell) when cell.DataType <> null && cell.DataType.HasValue -> 
               match cell.DataType.Value with
               | CellValues.Number -> typeof<decimal>
               | CellValues.SharedString -> typeof<string>
               | CellValues.Boolean -> typeof<bool>
               | CellValues.Date -> typeof<DateTime>
               | _ -> typeof<string>
            | _ as c -> typeof<string>

    interface IOfficeProvider with
       member x.GetFields() = 
           definedNames
           |> Map.toArray 
           |> Array.map (fun (n, addr) -> { FieldName = n; Type = (getCellType addr) })

       member x.ReadField(name:string) =
           let cell =
               let address = definedNames.[name]
               match tryGetCell address with
               | Some(cell) -> cell
               | None -> failwithf "Could not find cell(s) for range %s" name
           readCellValue cell

       member x.SetField(name:string, value:obj) =
           let cell =
               let address = definedNames.[name]
               match tryGetCell address with
               | Some(cell) -> cell
               | None -> failwithf "Could not find cell(s) for range %s" name
           writeCellValue cell value     
        
       member x.Commit(path) =
            doc.WorkbookPart.Workbook.Save()
            doc.Close()
            if File.Exists(path) then File.Delete(path)
            File.Copy(documentPath, path)
            (x :> IDisposable).Dispose()

       member x.Rollback() =
            doc.Close()
            (x :> IDisposable).Dispose()

       member x.Dispose() =
           doc.Dispose()
           if File.Exists(documentPath) && shadowCopy then File.Delete(documentPath) 
           
