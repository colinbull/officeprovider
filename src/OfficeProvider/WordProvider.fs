namespace OfficeProvider


type WordProvider(document:string) = 
     
     interface IOfficeProvider with
        member x.GetFields() = [||]

        member x.ReadField(name:string) = box (sprintf "Read word field %s from doc %s" name document)

        member x.SetField(name:string, value:obj) = ()

