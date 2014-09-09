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
        match Path.GetExtension(document) with
        | ".docx" -> (new WordProvider(document) :> IOfficeProvider) 
        | ".xlsx" -> (new ExcelProvider(resolutionPath, document) :> IOfficeProvider)
        | _ -> failwithf "Only docx (Word) and xlsx (Excel) files are currently supported"

    let staticParameters = [
        ProvidedStaticParameter("Document", typeof<string>)
        ProvidedStaticParameter("ResolutionPath", typeof<string>, "")
    ]

    do officeRootType.DefineStaticParameters(staticParameters, 
        fun typeName parameters ->
            let resolutionPath = (parameters.[1] :?> string)

            let serviceType = ProvidedTypeDefinition("DocumentTypes", None, HideObjectMethods = true)
            let documentType = ProvidedTypeDefinition("Document", Some typeof<IDisposable>, HideObjectMethods = true)

            use provider = createProviderInstance(resolutionPath,(parameters.[0] :?> string)) 

            provider.GetFields()
            |> Array.iter (fun field ->
                let fieldName = field.FieldName
                documentType.AddMember(ProvidedProperty(field.FieldName, field.Type, GetterCode = (fun args -> <@@ ((%%args.[0] : IDisposable) :?> IOfficeProvider).ReadField(fieldName) @@>)))
            )
            
            serviceType.AddMember(documentType)

            let rootType = ProvidedTypeDefinition(Assembly.LoadFrom config.RuntimeAssembly, rootNamespace, typeName, Some typeof<obj>, HideObjectMethods = true)
            
            rootType.AddMember(serviceType)
            rootType.AddMember(ProvidedMethod("Load", [ProvidedParameter("document", typeof<string>)], 
                                documentType, 
                                IsStaticMethod = true, 
                                InvokeCode = (fun args -> 
                                    <@@  new ExcelProvider("",(%%args.[0] : string)) :> IDisposable @@>)))

            rootType
    )


    do this.AddNamespace(rootNamespace, [officeRootType])


[<assembly:TypeProviderAssembly>]
do()

     
