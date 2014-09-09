namespace OfficeProvider

open System.IO
open System.Reflection
open Microsoft.FSharp.Core.CompilerServices
open ProviderImplementation.ProvidedTypes

[<TypeProvider>]
type OfficeTypeProvider(config) as this = 
    inherit TypeProviderForNamespaces()
    
    let rootNamespace = "OfficeProvider"
    let thisAssembly = Assembly.GetExecutingAssembly()
    let officeRootType = ProvidedTypeDefinition(thisAssembly, rootNamespace, "OfficeRoot", Some typeof<obj>)

    let createProviderInstance resolutionPath document = 
        match Path.GetExtension(document) with
        | ".docx" -> (new WordProvider(document) :> IOfficeProvider) 
        | ".xlsx" -> (new ExcelProvider(document) :> IOfficeProvider)
        | _ -> failwithf "Only docx (Word) and xlsx (Excel) files are currently supported"

    let staticParameters = [
        ProvidedStaticParameter("Document", typeof<string>)
        ProvidedStaticParameter("ResolutionPath", typeof<string>, "")
    ]

    do officeRootType.DefineStaticParameters(staticParameters, 
        fun typeName parameters ->
            let resolutionPath = (parameters.[1] :?> string)

            let ty = ProvidedTypeDefinition(thisAssembly, rootNamespace, typeName, Some typeof<obj>)
            let provider = createProviderInstance resolutionPath (parameters.[0] :?> string) 
            let documentType = ProvidedTypeDefinition("Document", Some typeof<obj>, HideObjectMethods = true)

            provider.GetFields()
            |> Array.iter (fun field -> 
                documentType.AddMember(ProvidedProperty(field.Name, field.Type, GetterCode = (fun args -> <@@ ((%%args.[0] : obj) :?> IOfficeProvider).ReadField(field.Name) @@>)))
            )
            
            ty.AddMember(documentType)
            ty.AddMember(ProvidedMethod("Load", [ProvidedParameter("document", typeof<string>)], documentType, IsStaticMethod = true, InvokeCode = (fun args -> <@@ obj() @@>)))

            ty
    )


    do this.AddNamespace("OfficeProvider", [officeRootType])


[<assembly:TypeProviderAssembly>]
do()

     
