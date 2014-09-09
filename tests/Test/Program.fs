// Learn more about F# at http://fsharp.net
// See the 'F# Tutorial' project for more help.

[<Literal>]
let path = @"D:\Proposal_Third_Party_Services.xlsx"

type Office = OfficeProvider.Office<path>

[<EntryPoint>]
let main argv = 
    
    let doc = Office.Load(path)
    doc.

    

    printfn "%A" argv
    0 // return an integer exit code
