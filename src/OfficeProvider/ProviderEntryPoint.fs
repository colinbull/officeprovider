namespace OfficeProvider

open System
open System.IO
open System.Reflection
open Microsoft.FSharp.Core.CompilerServices
open ProviderImplementation.ProvidedTypes

[<TypeProvider>]
type OfficeTypeProvider(config:TypeProviderConfig) as this = 
    inherit TypeProviderForNamespaces()
    
    let rootNamespace = "OfficeProvider"
    let thisAssembly = Assembly.GetExecutingAssembly()
    let officeRootType = ProvidedTypeDefinition(thisAssembly, rootNamespace, "Office", Some typeof<obj>)

    let createProviderInstance(resolutionPath,document) = 
        let fullPath = 
            if String.IsNullOrWhiteSpace(resolutionPath)
            then document
            else Path.Combine(resolutionPath, document)

        match Path.GetExtension(document) with
        | ".docx" -> (new WordProvider(fullPath) :> IOfficeProvider) 
        | ".xlsx" -> (new ExcelProvider(fullPath) :> IOfficeProvider)
        | _ -> failwithf "Only docx (Word) and xlsx (Excel) files are currently supported"

    let staticParameters = [
        ProvidedStaticParameter("Document", typeof<string>)
        ProvidedStaticParameter("ResolutionPath", typeof<string>, "")
    ]

    do officeRootType.DefineStaticParameters(staticParameters, 
        fun typeName parameters ->
            let resolutionPath = (parameters.[1] :?> string)

            let ty = ProvidedTypeDefinition(thisAssembly, rootNamespace, typeName, Some typeof<obj>)
            let provider = createProviderInstance(resolutionPath,(parameters.[0] :?> string)) 
            let documentType = ProvidedTypeDefinition("Document", None, HideObjectMethods = true)

            provider.GetFields()
            |> Array.iter (fun field -> 
                documentType.AddMember(ProvidedProperty(field.FieldName, field.Type, GetterCode = (fun args -> <@@ ((%%args.[0] : obj) :?> IOfficeProvider).ReadField(field.FieldName) @@>)))
            )
            
            ty.AddMember(documentType)
            ty.AddMember(ProvidedMethod("Load", [ProvidedParameter("document", typeof<string>)], documentType, IsStaticMethod = true, InvokeCode = (fun args -> <@@  createProviderInstance("",(%%args.[0] : string)) @@>)))

            ty
    )


    do this.AddNamespace("OfficeProvider", [officeRootType])


[<assembly:TypeProviderAssembly>]
do()

     
