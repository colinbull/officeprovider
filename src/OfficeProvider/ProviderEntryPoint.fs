namespace OfficeProvider

open System
open System.IO
open System.Reflection
open Microsoft.FSharp.Core.CompilerServices
open ProviderImplementation.ProvidedTypes
open Microsoft.FSharp.Quotations

type OfficeTypeProvider(config:TypeProviderConfig, 
                        rootTypeCtor: (Assembly * string -> ProvidedTypeDefinition), 
                        providerCtor : (string * string * bool -> IOfficeProvider),
                        loadExpr : (string * bool * Expr list) -> Expr) as this = 
    inherit TypeProviderForNamespaces()
    
    let rootNamespace = "OfficeProvider"
    let thisAssembly = Assembly.GetExecutingAssembly()
    let officeRootType = rootTypeCtor (thisAssembly, rootNamespace)


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

            let provider = providerCtor(resolutionPath, documentPath, shadowCopy) 

            provider.GetFields()
            |> Array.iter (fun field ->
                let fieldName = field.FieldName
                documentType.AddMember(ProvidedProperty(field.FieldName, field.Type, 
                                        GetterCode = (fun args -> <@@ ((%%args.[0] : ITransacted) :?> IOfficeProvider).ReadField(fieldName) @@>),
                                        SetterCode = (fun args -> <@@ ((%%args.[0] : ITransacted) :?> IOfficeProvider).SetField(fieldName, %%Expr.Coerce(args.[1],typeof<obj>)) @@>)))
            )

            provider.Dispose()

            serviceType.AddMember(documentType)

            let rootType = ProvidedTypeDefinition(Assembly.LoadFrom config.RuntimeAssembly, rootNamespace, typeName, Some typeof<obj>, HideObjectMethods = true)
            
            rootType.AddMember(serviceType)
            rootType.AddMember(ProvidedMethod("Load", [ProvidedParameter("document", typeof<string>)], 
                                documentType, 
                                IsStaticMethod = true, 
                                InvokeCode = (fun args -> loadExpr(resolutionPath, shadowCopy, args))))

            rootType
    )
    
    do this.AddNamespace(rootNamespace, [officeRootType])

[<TypeProvider>]
type ExcelTypeProvider(config:TypeProviderConfig) =
    inherit OfficeTypeProvider(
        config, 
        (fun (assm, ns) -> ProvidedTypeDefinition(assm, ns, "Excel", Some typeof<obj>)),
        (fun (resPath, filePath, shadowCopy) -> new ExcelProvider(resPath, filePath, shadowCopy) :> IOfficeProvider),
        (fun (resPath, shadowCopy, args) -> 
            <@@  
                 let doc = (%%args.[0] : string)
                 new ExcelProvider(resPath, doc, shadowCopy) :> ITransacted @@>)   
    )

[<TypeProvider>]
type WordTypeProvider(config:TypeProviderConfig) =
    inherit OfficeTypeProvider(
        config, 
        (fun (assm, ns) -> ProvidedTypeDefinition(assm, ns, "Word", Some typeof<obj>)),
        (fun (resPath, filePath, shadowCopy) -> new WordProvider(resPath, filePath, shadowCopy) :> IOfficeProvider),
        (fun (resPath, shadowCopy, args) -> 
            <@@  
                 let doc = (%%args.[0] : string)
                 new WordProvider(resPath, doc, shadowCopy) :> ITransacted @@>)    
    )

[<assembly:TypeProviderAssembly>]
do()

     
