namespace OfficeProvider

module Inference =
        
    open System
    open System.Globalization       
    open ProviderImplementation.ProvidedTypes
    open Microsoft.FSharp.Quotations
    open OfficeProvider
                
    let defaultCulture = CultureInfo.CurrentCulture
    let nullValues = ["N/A"; "-"; ""; "NULL"]
    let isNullValue str =
        if String.IsNullOrWhiteSpace(str)
        then true
        else nullValues |> List.exists ((=) (str.ToUpper())) 
           
    let inferPrimitive (cultureInfo:CultureInfo) (str:string) =
        let (|Parse|_|) f v =
            match f str with
            | true, v -> Some v
            | false, _ -> None

        match str with
        | Parse Int32.TryParse v when v = 0 -> Primitive (typeof<Bool0>, false)
        | Parse Int32.TryParse v when v = 1 -> Primitive (typeof<Bool1>, false)
        | Parse Boolean.TryParse _ -> Primitive (typeof<bool>, false)
        | Parse Int32.TryParse _ -> Primitive (typeof<Int32>, false)
        | Parse Int64.TryParse _ -> Primitive (typeof<Int64>, false)
        | Parse Decimal.TryParse _ -> Primitive (typeof<decimal>, false)
        | Parse Double.TryParse _ -> Primitive (typeof<float>, false)
        | Parse Guid.TryParse _ -> Primitive (typeof<Guid>, false)
        | Parse DateTime.TryParse _ -> Primitive (typeof<DateTime>, false)
        | _ when isNullValue str -> InferredType.Null
        | _ -> Primitive (typeof<string>, false)


    let private unifactionRules =
        [
            typeof<bool>, [typeof<Bool0>; typeof<Bool1>; typeof<bool>]
            typeof<Int32>, [typeof<Bool0>; typeof<Bool1>; typeof<bool>; typeof<Int32>]
            typeof<Int64>, [typeof<Bool0>; typeof<Bool1>; typeof<bool>; typeof<Int32>; typeof<Int64>]
            typeof<float>, [typeof<Bool0>; typeof<Bool1>; typeof<bool>; typeof<Int32>; typeof<Int64>; typeof<float>]
            typeof<decimal>, [typeof<Bool0>; typeof<Bool1>; typeof<bool>; typeof<Int32>; typeof<Int64>; typeof<float>; typeof<decimal>]
        ]

    let unify (values:seq<InferredType>) =

        let unifyPrimitive optional typ1 typ2 =
            let tryFindCommonType typ (super,sub) =
                if typ = super || (sub |> List.exists ((=) typ))
                then Some (super, sub)
                else None

            let unifyType typ1 typ2 =
                unifactionRules
                |> List.tryPick (tryFindCommonType typ1 >> Option.bind (tryFindCommonType typ2))

            match unifyType typ1 typ2 with
            | Some (t, _) -> Primitive (t, optional)
            | None ->
                match unifyType typ2 typ1 with
                | Some (t,_) -> Primitive (t, optional)
                | None -> Primitive (typeof<string>, optional)

        let unifyType typ1 typ2 =
            match typ1, typ2 with
            | Primitive (t1, o1), Primitive (t2, o2) -> unifyPrimitive (o1 || o2) t1 t2
            | InferredType.Null, (Primitive (t, o1)) -> Primitive(t, (true || o1))
            | Primitive (t,o1), InferredType.Null -> Primitive(t, (o1 || true))
            | InferredType.Null, InferredType.Null -> Null
            | _, _ -> failwithf "Unable to unify type %A %A" typ1 typ2

        values
        |> Seq.fold (fun s x -> unifyType s x) (Seq.head values)

    let inferPrimitiveType culture  (values : seq<string>) =
        values
        |> Seq.map (inferPrimitive culture)
        |> unify
