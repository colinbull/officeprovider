﻿namespace OfficeProvider

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
    
    type Field = {
        FieldName : string
        Type : Type
    }

    type ITransacted = 
        inherit IDisposable
        abstract Commit : string -> unit
        abstract Rollback : unit -> unit

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

        let firstOrCreate (ctor : unit -> 'a) (f : 'a -> 'b) (e:#OpenXmlElement) = 
            match e.Elements<'a>() |> Seq.tryHead with
            | Some(r) -> f r
            | None -> let r = ctor() in e.Append(r); f r
               
