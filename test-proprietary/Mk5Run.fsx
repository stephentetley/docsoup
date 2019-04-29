// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"
#r "System.Xml.Linq"
#r "System.IO.FileSystem.Primitives"

open System.IO
open System.Text.RegularExpressions

#I @"C:\Users\stephen\.nuget\packages\DocumentFormat.OpenXml\2.9.1\lib\netstandard1.3"
#r "DocumentFormat.OpenXml"
#I @"C:\Users\stephen\.nuget\packages\system.io.packaging\4.5.0\lib\netstandard1.3"
#r "System.IO.Packaging"


// Use FSharp.Data for CSV output
#I @"C:\Users\stephen\.nuget\packages\FSharp.Data\3.1.1\lib\netstandard2.0"
#I @"C:\Users\stephen\.nuget\packages\FSharp.Data\3.1.1\typeproviders\fsharp41\netstandard2.0"
#r @"FSharp.Data.dll"
#r @"FSharp.Data.DesignTime"
open FSharp.Data

#load @"..\src\DocSoup\Internal\Common.fs"
#load @"..\src\DocSoup\Internal\OpenXml.fs"
#load @"..\src\DocSoup\ExtractMonad.fs"
#load @"..\src\DocSoup\CellExtract.fs"
#load @"..\src\DocSoup\RowExtract.fs"
#load @"..\src\DocSoup\TableExtract.fs"
#load @"..\src\DocSoup\BodyExtract.fs"
#load @"..\src\DocSoup\DocumentExtract.fs"
open DocSoup

#load @"Mk5ReplacementExtract.fs"
open Mk5ReplacementExtract


let localFile (fileName:string) : string = 
    System.IO.Path.Combine (__SOURCE_DIRECTORY__ , "../data", fileName)


let getOk (ans:Result<'a, ErrMsg>) : seq<'a> = 
    match ans with
    | Ok a -> Seq.singleton a
    | Error msg -> printfn "%s" msg; Seq.empty


let sampleFile = @"G:\work\Projects\rtu\mk5-mmims\SAMPLE Upgrade Site Works.docx"

let demo01 () = 
    runDocumentExtractor sampleFile (body &>> extractSiteInfo)


let demo02 () = 
    runDocumentExtractor sampleFile 
        (body &>> tupleM2 extractSiteInfo extractInstallInfo)

