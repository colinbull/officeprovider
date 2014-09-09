namespace OfficeProvider

open System
open System.IO
open System.Linq
open System.Text.RegularExpressions
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing

type WordProvider(resolutionPath:string, document:string) = 
     let documentPath = 
            if String.IsNullOrWhiteSpace(resolutionPath)
            then document
            else Path.Combine(resolutionPath, document)

     let doc =
        if File.Exists(documentPath)
        then WordprocessingDocument.Open(documentPath, true, new OpenSettings(AutoSave = true))
        else raise(FileNotFoundException("Could not find file", documentPath)) 

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
        |> Map.ofArray
           
     interface IOfficeProvider with
        member x.GetFields() =
            contentControls |> Map.toArray |> Array.map (fun (name, _) -> { FieldName = name; Type = typeof<String> })

        member x.ReadField(name:string) =
            contentControls.[name].Descendants<Text>().Single().Text |> box


        member x.SetField(name:string, value:obj) = ()

        member x.Dispose() = doc.Close()

