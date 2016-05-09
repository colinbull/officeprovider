namespace OfficeProvider

open System
open System.IO
open System.ComponentModel
open DocumentFormat.OpenXml

module File = 
    
    let getPath resolutionPath path extension shadowCopy = 
        let originalPath = 
            if String.IsNullOrWhiteSpace(resolutionPath)
            then new FileInfo(path)
            else new FileInfo(Path.Combine(resolutionPath, path))

        if originalPath.Exists
        then 
            if shadowCopy
            then 
                let tempPath = Path.GetTempFileName()
                let xlsxTemp = Path.ChangeExtension(tempPath, extension)

                if tempPath = xlsxTemp then File.Delete(tempPath)

                File.Copy(originalPath.FullName, xlsxTemp, true)
                xlsxTemp
            else originalPath.FullName
        else raise(FileNotFoundException("Could not find file", originalPath.FullName))  

[<AutoOpen>]
module Types =

    type Bool0 = Bool0
    type Bool1 = Bool1

    type InferredType =
         | Primitive of Type * optional:bool
         | Record of string * (string * InferredType) list
         | Null
         override x.ToString() =
            match x with
            | Primitive(t, opt) -> sprintf "Primitive(%A, %b)" t opt
            | Null -> "NULL"
            | Record(name, fields) -> sprintf "%s {%s}" name (String.Join(";", fields |> List.map(fun (name, t) -> sprintf "%s = %A" name t )))
                 
    type Field = {
        FieldName : string
        Type : InferredType
    }

    type ITransacted = 
        inherit IDisposable
        abstract Commit : string -> unit
        abstract Rollback : unit -> unit
    
    type ProviderInitParameters = {
        ResolutionPath : string
        DocumentPath : string
        ShadowCopy : bool
        AllowNameEquality : bool
    }

    type IOfficeProvider =
        inherit ITransacted
        //Gets the fields that can be used to fill in types
        //In word these are actually fields in excel these
        //are named ranges.
        abstract GetFields : unit -> Field[]

        //Gets the value of the field
        abstract ReadField : string -> obj

        //Sets the value of the field
        abstract SetField : string * obj -> unit

[<CompilationRepresentation(CompilationRepresentationFlags.ModuleSuffix)>]
module Field = 

    open System
    open Microsoft.FSharp.Quotations
    open ProviderImplementation.ProvidedTypes
    
    let rec toProvidedProperty (serviceTypes:ProvidedTypeDefinition) (container:ProvidedTypeDefinition) (field:Field) =
        let makeOptional t =
            let t = typedefof<Nullable<_>>
            t.MakeGenericType([|t|]) 

        let getter fieldName =
            (fun (args:Expr list) -> <@@ ((%%args.[0] : ITransacted) :?> IOfficeProvider).ReadField(fieldName) @@>)
        let setter fieldName = 
            (fun (args:Expr list) -> <@@ ((%%args.[0] : ITransacted) :?> IOfficeProvider).SetField(fieldName, Expr.Coerce(args.[1],typeof<obj>)) @@>)
        
        match field.Type with
        | Primitive (t, false) ->
            ProvidedProperty(field.FieldName, t, GetterCode = (getter field.FieldName), SetterCode = (setter field.FieldName))
        | Primitive (t, true) ->
            let t = makeOptional t
            ProvidedProperty(field.FieldName, t, GetterCode = (getter field.FieldName), SetterCode = (setter field.FieldName))
        | Null ->
            let t = makeOptional typeof<string>
            ProvidedProperty(field.FieldName, t, GetterCode = (getter field.FieldName), SetterCode = (setter field.FieldName))
        | Record(name,fields) ->
            let fields =
                [
                    for (n,f) in fields do
                        yield toProvidedProperty serviceTypes container { FieldName = n; Type = f }
                ]
            let providedT = ProvidedTypeDefinition(name, None)
            providedT.AddMembers(fields)
            serviceTypes.AddMember(providedT)
            ProvidedProperty(field.FieldName, providedT, GetterCode = (fun args -> <@@ obj() @@>))



[<AutoOpen>]
module Helpers = 
    
    type MaybeBuilder() = 
        member __.Bind(m,f) = Option.bind f m
        member __.Return(x) = Some x
        member __.ReturnFrom(x) = x

    let maybe = MaybeBuilder()

    module Seq = 
        
        let tryHead (source : seq<_>) = 
            use e = source.GetEnumerator()
            if e.MoveNext()
            then Some(e.Current)
            else None //empty list   

        let tryHeadOrCreate ctor (source : _ seq) = 
            match tryHead source with
            | Some(f) when f <> null -> f
            | _ -> ctor()
        
        let ofObject (v:obj) = 
            seq {
                if v <> null
                then
                    for prop in TypeDescriptor.GetProperties(v.GetType()) do
                        let value = prop.GetValue(v)
                        if value <> null
                        then yield value.ToString()
                        else yield String.Empty
            }
                
    module Xml = 

        open DocumentFormat.OpenXml
        open DocumentFormat.OpenXml.Packaging
        open DocumentFormat.OpenXml.Wordprocessing

        let firstOrCreate (ctor : unit -> 'a) (f : 'a -> 'b) (e:#OpenXmlElement) = 
            match e.Descendants<'a>() |> Seq.tryHead with
            | Some(r) -> f r
            | None -> let r = ctor() in e.Append(r); f r
        
        let innerTextConcat join (e:seq<#OpenXmlElement>) = 
            String.Join(join, e |> Seq.map (fun x -> x.InnerText))

        let rec getText (e:OpenXmlElement) = 
            e.Descendants() 
            |> Seq.map (fun x -> 
                match box x with
                | :? SdtRun as r ->  r.Descendants<Text>() |> innerTextConcat " "
                | :? SdtBlock as r -> String.Join(Environment.NewLine, getText (r :> OpenXmlElement))
                | :? Paragraph as r -> String.Join(Environment.NewLine, getText (r :> OpenXmlElement))
                | _ -> failwithf "Failed to get text on %A" x
            )
