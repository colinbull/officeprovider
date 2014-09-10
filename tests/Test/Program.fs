// Learn more about F# at http://fsharp.net
// See the 'F# Tutorial' project for more help.

[<Literal>]
let path = @"D:\Appdev\officeprovider\docs\content\SimpleInvoice.xlsx"

[<Literal>]
let wordPath = @"D:\Appdev\officeprovider\docs\content\Billing statement.docx"

type Excel = OfficeProvider.Office<path>

//type Word = OfficeProvider.Office<wordPath>

[<EntryPoint>]
let main argv = 
    
//    use wordDoc = Word.Load(wordPath)
//
//    printfn "%s" wordDoc.Date

    use doc = Excel.Load(path)
    
    doc.Name <- "Colin Bull"
    printfn "%s" doc.Name

    System.Console.ReadLine() |> ignore
    0 // return an integer exit code
