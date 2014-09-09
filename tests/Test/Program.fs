// Learn more about F# at http://fsharp.net
// See the 'F# Tutorial' project for more help.

[<Literal>]
let path = @"D:\Appdev\officeprovider\docs\content\SimpleInvoice.xlsx"

type Office = OfficeProvider.Office<path>

[<EntryPoint>]
let main argv = 
    
    use doc = Office.Load(path)
    printfn "%s" doc.Name

    System.Console.ReadLine() |> ignore
    0 // return an integer exit code
