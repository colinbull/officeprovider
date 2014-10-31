namespace OfficeProvider

open System
open System.IO
open System.Linq
open System.Collections
open System.ComponentModel
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing

type WordProvider(resolutionPath:string, document:string, shadowCopy:bool) = 
     let documentPath = File.getPath resolutionPath document "docx" shadowCopy
     let doc = WordprocessingDocument.Open(documentPath, true)

     let contentControls =
        [|
            for cc in doc.MainDocumentPart.Document.Descendants<SdtElement>() do
                yield cc;
            for header in doc.MainDocumentPart.HeaderParts do
                for cc in header.Header.Descendants<SdtElement>() do
                    yield cc;
            for footer in doc.MainDocumentPart.FooterParts do
                for  cc in footer.Footer.Descendants<SdtElement>() do
                    yield cc;
            if (doc.MainDocumentPart.FootnotesPart <> null)
            then
                for  cc in doc.MainDocumentPart.FootnotesPart.Footnotes.Descendants<SdtElement>() do
                    yield cc;
            if (doc.MainDocumentPart.EndnotesPart <> null)
            then
                for cc in doc.MainDocumentPart.EndnotesPart.Endnotes.Descendants<SdtElement>() do
                    yield cc
        |] 
        |> Array.map (fun cc -> cc.SdtProperties.GetFirstChild<Tag>().Val.Value, cc)
        |> Seq.groupBy fst |> Seq.map (fun (k,v) -> k, v |> Seq.map snd |> Seq.toArray)
        |> Map.ofSeq

     let writeTable (values: _ []) (target:SdtElement) = 
        let props = 
            new TableProperties(
                new TableBorders
                    ([|
                        new TopBorder(Val = new EnumValue<BorderValues>(BorderValues.Single), Size = new UInt32Value(12u)) :> OpenXmlElement
                        new BottomBorder(Val = new EnumValue<BorderValues>(BorderValues.Single), Size = new UInt32Value(12u)) :> OpenXmlElement
                        new LeftBorder(Val = new EnumValue<BorderValues>(BorderValues.Single), Size = new UInt32Value(12u)) :> OpenXmlElement
                        new RightBorder(Val = new EnumValue<BorderValues>(BorderValues.Single), Size = new UInt32Value(12u)) :> OpenXmlElement
                        new InsideHorizontalBorder(Val = new EnumValue<BorderValues>(BorderValues.Single), Size = new UInt32Value(12u)) :> OpenXmlElement
                        new InsideVerticalBorder(Val = new EnumValue<BorderValues>(BorderValues.Single), Size = new UInt32Value(12u)) :> OpenXmlElement
                    |]))
        let table = new Table()
        table.Append(props)

        let addRow (table:Table) (row:obj) = 
            if row <> null
            then
                let tr = new TableRow()
                let cells = 
                    Seq.ofObject row 
                    |> Seq.map (fun v -> new TableCell([|new Paragraph(new Run(new Text(v))) :> OpenXmlElement|]) :> OpenXmlElement)
                tr.Append(cells)
                table.Append(tr)
            table
        
        target.Append(values |> Array.fold addRow table)

     let writeString (value:string) (target:SdtElement) = 
         target
         |> Xml.firstOrCreate (fun () -> Paragraph()) id
         |> Xml.firstOrCreate (fun () -> Run()) id
         |> Xml.firstOrCreate (fun () -> Text()) (fun (a : Text) -> a.Text <- value)
           
     interface IOfficeProvider with
       member x.GetFields() =
           contentControls 
           |> Map.toArray 
           |> Array.map (fun (name, vs) -> { FieldName = name; Type = typeof<String> })

       member x.ReadField(name:string) =
           match contentControls.[name] with
           | [|t|] -> t.Descendants<Text>().Single().Text |> box
           | ts -> ts |> Array.map (fun t -> t.Descendants<Text>().Single().Text |> box) |> box

       member x.SetField(name:string, value:obj) =
            let target = contentControls.[name]

            let setContent (target:SdtElement) = 
                let typ = value.GetType()
                if typ.IsArray
                then writeTable ((value :?> IEnumerable).Cast<obj>().ToArray()) target
                else writeString (value.ToString()) target

            match target with
            | [|t|] -> setContent t
            | ts -> ts |> Array.iter setContent

       member x.Commit(path) = 
            if File.Exists(path) then File.Delete(path)
            File.Copy(documentPath, path)
            (x :> IDisposable).Dispose()

       member x.Rollback() = 
            (x :> IDisposable).Dispose()

       member x.Dispose() =
           doc.Dispose()
           if File.Exists(documentPath) && shadowCopy then File.Delete(documentPath) 
          

