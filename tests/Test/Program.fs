// Learn more about F# at http://fsharp.net
// See the 'F# Tutorial' project for more help.

[<Literal>]
let path = @"D:\Appdev\officeprovider\docs\content\SimpleInvoice.xlsx"

[<Literal>]
let wordPath = @"D:\Appdev\officeprovider\docs\content\Billing statement.docx"


type Excel = OfficeProvider.Excel<path>
type Word = OfficeProvider.Word<wordPath>

[<EntryPoint>]
let main argv = 
    
    use doc = Excel.Load(path)

    doc.Name <- "My Company"
    
    printfn "%s" doc.Name

    doc.Commit(@"D:\Appdev\officeprovider\docs\content\SimpleInvoice_Updated.xlsx")

    use word = Word.Load(wordPath)
    
    printfn "%s" word.Company
    word.Company <- "My Company"
    
    printfn "%s" word.Company 
    
    word.Commit(@"D:\Appdev\officeprovider\docs\content\BillingStatement_updated.docx")

    System.Console.ReadLine() |> ignore
    0 // return an integer exit code
