namespace OfficeProvider

open System

[<AutoOpen>]
module Types =
    
    type Field = {
        Name : string
        Type : Type
    }

    type IOfficeProvider = 
        //Gets the fields that can be used to fill in types
        //In word these are actually fields in excel these
        //are named ranges.
        abstract GetFields : unit -> Field[]

        //Gets the value of the field
        abstract ReadField : string -> obj

        //Sets the value of the field
        abstract SetField : string * obj -> unit