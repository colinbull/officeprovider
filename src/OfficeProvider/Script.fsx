// Learn more about F# at http://fsharp.net. See the 'F# Tutorial' project
// for more guidance on F# programming.

open System
open System.Globalization       

fsi.AddPrinter(fun (x:Type) -> x.FullName)

type Bool0 = Bool0
type Bool1 = Bool1
        
type InferredType =
     | Primitive of Type * optional:bool
     | Null
    
let defaultCulture = CultureInfo.CurrentCulture

let isNullValue str =
    if String.IsNullOrWhiteSpace(str)
    then true
    else 
        match str.ToUpper() with
        | "N/A"
        | "-"
        | "NULL" -> true
        | _ -> false

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
        typeof<string>, [] //Hhm!! Maybe not!!
        typeof<DateTime>, []
        typeof<Guid>, []
        typeof<bool>, [typeof<Bool0>; typeof<Bool1>]
        typeof<int>,  [typeof<Bool0>; typeof<Bool1>; typeof<Int32>]
        typeof<int64>, [typeof<Bool0>; typeof<Bool1>; typeof<Int32>; typeof<Int64>]
        typeof<float>, [typeof<Bool0>; typeof<Bool1>; typeof<Int32>; typeof<Int64>; typeof<float>]
        typeof<decimal>, [typeof<Bool0>; typeof<Bool1>; typeof<Int32>; typeof<Int64>; typeof<float>; typeof<decimal>]
    ]
    
let unify (values:seq<InferredType>) =

    let unifyPrimitive optional typ1 typ2 =
        let tryFindCommonType typ (super,sub) =
            if typ = super || (sub |> List.exists ((=) typ))
            then Some (super,sub)
            else None

        let unifyType typ1 typ2 =
            unifactionRules
            |> List.tryPick (tryFindCommonType typ1)
            |> Option.bind (tryFindCommonType typ2)

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
          
    values
    |> Seq.fold (fun s x -> unifyType s x) (Seq.head values)

let inferPrimitiveType (values : seq<string>) =
    values
    |> Seq.map (inferPrimitive defaultCulture)
    |> unify
