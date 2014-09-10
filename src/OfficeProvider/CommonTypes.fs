namespace OfficeProvider

open System

module Seq = 
    
    let findIndexOrMax p (source:seq<_>) = 
            if source = null then raise(NullReferenceException())
            use ie = source.GetEnumerator() 
            let rec loop i = 
                if ie.MoveNext() then 
                    if p ie.Current then Choice1Of2 i
                    else loop (i+1)
                else
                    Choice2Of2 i
            loop 0

[<AutoOpen>]
module Types =
    
    type Field = {
        FieldName : string
        Type : Type
    }

    type IOfficeProvider =
        inherit IDisposable
        //Gets the fields that can be used to fill in types
        //In word these are actually fields in excel these
        //are named ranges.
        abstract GetFields : unit -> Field[]

        //Gets the value of the field
        abstract ReadField : string -> obj

        //Sets the value of the field
        abstract SetField : string * obj -> unit