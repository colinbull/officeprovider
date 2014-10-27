namespace OfficeProvider

open System
open System.IO
open System.Reflection
open Microsoft.FSharp.Core.CompilerServices
open ProviderImplementation.ProvidedTypes
open Microsoft.FSharp.Quotations

[<TypeProvider>]
type OfficeTypeProvider(config:TypeProviderConfig) as this = 
    inherit TypeProviderForNamespaces()
    
    let rootNamespace = "OfficeProvider"
    let thisAssembly = Assembly.GetExecutingAssembly()
    let officeRootType = ProvidedTypeDefinition(thisAssembly, rootNamespace, "Office", Some typeof<obj>)

    let createProviderInstance(resolutionPath,document) = 
        match Path.GetExtension(document) with
        | ".docx" -> (new WordProvider(resolutionPath, document, true) :> IOfficeProvider) 
        | ".xlsx" -> (new ExcelProvider(resolutionPath, document, true) :> IOfficeProvider)
        | _ -> failwithf "Only docx (Word) and xlsx (Excel) files are currently supported"

    let staticParameters = [
        ProvidedStaticParameter("Document", typeof<string>)
        ProvidedStaticParameter("WorkingDirectory", typeof<string>, "")
        ProvidedStaticParameter("CopySourceFile", typeof<bool>, true)
    ]

    do officeRootType.DefineStaticParameters(staticParameters, 
        fun typeName parameters ->
            let documentPath = (parameters.[0] :?> string)
            let resolutionPath = 
                let respath = (parameters.[1] :?> string)
                if String.IsNullOrWhiteSpace respath 
                then config.ResolutionFolder
                else respath
            let shadowCopy = (parameters.[2] :?> bool)

            let serviceType = ProvidedTypeDefinition("DocumentTypes", None, HideObjectMethods = true)
            let documentType = ProvidedTypeDefinition("Document", Some typeof<ITransacted>, HideObjectMethods = true)

            use provider = createProviderInstance(resolutionPath, documentPath) 

            provider.GetFields()
            |> Array.iter (fun field ->
                let fieldName = field.FieldName
                documentType.AddMember(ProvidedProperty(field.FieldName, field.Type, 
                                        GetterCode = (fun args -> <@@ ((%%args.[0] : ITransacted) :?> IOfficeProvider).ReadField(fieldName) @@>),
                                        SetterCode = (fun args -> <@@ ((%%args.[0] : ITransacted) :?> IOfficeProvider).SetField(fieldName, %%Expr.Coerce(args.[1],typeof<obj>)) @@>)))
            )
            
            serviceType.AddMember(documentType)

            let rootType = ProvidedTypeDefinition(Assembly.LoadFrom config.RuntimeAssembly, rootNamespace, typeName, Some typeof<obj>, HideObjectMethods = true)
            
            rootType.AddMember(serviceType)
            rootType.AddMember(ProvidedMethod("Load", [ProvidedParameter("document", typeof<string>)], 
                                documentType, 
                                IsStaticMethod = true, 
                                InvokeCode = (fun args -> 
                                    <@@  
                                        let doc = (%%args.[0] : string)
                                        if doc.EndsWith("xlsx")
                                        then new ExcelProvider(resolutionPath, doc, shadowCopy) :> ITransacted
                                        else new WordProvider(resolutionPath, doc, shadowCopy) :> ITransacted @@>)))

            rootType
    )
    
    do this.AddNamespace(rootNamespace, [officeRootType])


[<assembly:TypeProviderAssembly>]
do()

     
