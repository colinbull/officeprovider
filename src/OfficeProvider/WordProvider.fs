namespace OfficeProvider

open System
open System.IO
open System.Linq
open System.Collections
open System.ComponentModel
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing

type WordProvider(parameters:ProviderInitParameters) = 
     let documentPath = File.getPath parameters.ResolutionPath parameters.DocumentPath "docx" parameters.ShadowCopy
     let doc = WordprocessingDocument.Open(documentPath, true, new OpenSettings(AutoSave = true))

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
        |> Seq.groupBy (fun cc -> cc.SdtProperties.GetFirstChild<Tag>().Val.Value)
        |> Seq.collect (fun (key, elems) -> 
            if (Seq.length elems) = 1
            then seq { yield key, Seq.head elems }
            else Seq.mapi (fun i e -> if i = 0 then key, e else key + (string i), e) elems
        )
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
           |> Array.map (fun (name, _) -> { FieldName = name; Type = typeof<String> })

       member x.ReadField(name:string) =
           contentControls.[name].Descendants<Text>().Single().Text |> box
           

       member x.SetField(name:string, value:obj) =
            let target = contentControls.[name]

            let setContent (target:SdtElement) = 
                let typ = value.GetType()
                if typ.IsArray
                then writeTable ((value :?> IEnumerable).Cast<obj>().ToArray()) target
                else writeString (value.ToString()) target

            setContent target

       member x.Commit(path) =
            doc.MainDocumentPart.Document.Save()
            doc.Close()
            if File.Exists(path) then File.Delete(path)
            File.Copy(documentPath, path)
            (x :> IDisposable).Dispose()

       member x.Rollback() = 
            (x :> IDisposable).Dispose()

       member x.Dispose() =
           doc.Dispose()
           if File.Exists(documentPath) && parameters.ShadowCopy then File.Delete(documentPath) 
          

