// Learn more about F# at http://fsharp.net
// See the 'F# Tutorial' project for more help.

open System
open OfficeProvider

[<Literal>]
let path = @"/Users/colinbull/appdev/officeprovider/docs/content/SimpleInvoice.xlsx"

[<Literal>]
let wordPath = @"/Users/colinbull/appdev/officeprovider/docs/content/Billing statement.docx"


type Excel = OfficeProvider.Excel<path>
//type Word = OfficeProvider.Word<wordPath>

[<EntryPoint>]
let main argv = 
    
    use doc = Excel.Load(path)

    //doc.DueDate
    //doc.InvoiceNumber <- 1123
    //doc.Name <- "My Company"
    //doc.Address2 <- DateTime.Now
        
    printfn "%A" doc.Name

    doc.Commit(@"/Users/colinbull/appdev/officeprovider/docs/content/SimpleInvoice_Updated.xlsx")

    // use word = Word.Load(wordPath)
    
    // printfn "%s" word.Address
    // word.Company <- "My Company"
    
    // printfn "%s" word.Company 
    
    // word.Commit(@"/Users/colinbull/appdev/officeprovider/docs/content/BillingStatement_updated.docx")

    System.Console.ReadLine() |> ignore
    0 // return an integer exit code
