namespace System
open System.Reflection

[<assembly: AssemblyTitleAttribute("OfficeProvider")>]
[<assembly: AssemblyProductAttribute("OfficeProvider")>]
[<assembly: AssemblyDescriptionAttribute("An office type provider for xlsx and docx files")>]
[<assembly: AssemblyVersionAttribute("1.0")>]
[<assembly: AssemblyFileVersionAttribute("1.0")>]
do ()

module internal AssemblyVersionInformation =
    let [<Literal>] Version = "1.0"
    let [<Literal>] InformationalVersion = "1.0"
