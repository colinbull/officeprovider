namespace OfficeProvider

open System

type ExcelProvider(document:string) = 
     
     interface IOfficeProvider with
        member x.GetFields() = [|
            { Name = "Field A"; Type = typeof<string> }
            { Name = "Field B"; Type = typeof<DateTime> }
        |]

        member x.ReadField(name:string) = box (sprintf "Read word field %s from doc %s" name document)

        member x.SetField(name:string, value:obj) = ()
