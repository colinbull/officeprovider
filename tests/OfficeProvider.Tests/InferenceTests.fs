module InferenceTests

open System
open System.Globalization
open NUnit.Framework
open OfficeProvider

let enCulture = CultureInfo.GetCultureInfo("en-GB")  

[<Test>]
let ``should infer type as int``() =
    let actual = Inference.inferPrimitive enCulture "1123" 
    let expected = InferredType.Primitive(typeof<int32>, false)
    Assert.AreEqual(expected, actual)

[<Test>]
let ``should infer type as bool``() =
    let actual = Inference.inferPrimitive enCulture "0"
    let expected = InferredType.Primitive(typeof<Bool0>, false)
    Assert.AreEqual(expected, actual)

[<Test>]
let ``should infer type as bool1``() =
    let actual = Inference.inferPrimitive enCulture "1"
    let expected = InferredType.Primitive(typeof<Bool1>, false)
    Assert.AreEqual(expected, actual)

[<Test>]
let ``should unify to bool``() =
     let inferredTypes = [
            InferredType.Primitive(typeof<Bool0>,false)
            InferredType.Primitive(typeof<Bool1>,false)
         ]

     let actual = Inference.unify inferredTypes
     let expected = InferredType.Primitive(typeof<bool>, false)
     Assert.AreEqual(expected, actual)

[<Test>]
let ``should unify to int``() =
     let inferredTypes = [
            InferredType.Primitive(typeof<Bool0>,false)
            InferredType.Primitive(typeof<Bool1>,false)
            InferredType.Primitive(typeof<Int32>,false)
         ]

     let actual = Inference.unify inferredTypes
     let expected = InferredType.Primitive(typeof<int32>, false)
     Assert.AreEqual(expected, actual)

[<Test>]
let ``should unify to int64``() =
     let inferredTypes = [
            InferredType.Primitive(typeof<Bool0>,false)
            InferredType.Primitive(typeof<Bool1>,false)
            InferredType.Primitive(typeof<Int64>,false)
         ]

     let actual = Inference.unify inferredTypes
     let expected = InferredType.Primitive(typeof<int64>, false)
     Assert.AreEqual(expected, actual)



     
        

