namespace OfficeProvider

open System
open System.IO

module File = 
    
    let getPath resolutionPath path shadowCopy = 
        let originalPath = 
            if String.IsNullOrWhiteSpace(resolutionPath)
            then new FileInfo(path)
            else new FileInfo(Path.Combine(resolutionPath, path))

        if originalPath.Exists
        then 
            if shadowCopy
            then 
                let tempPath = Path.GetTempFileName()
                let xlsxTemp = Path.ChangeExtension(tempPath, "xlsx")

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