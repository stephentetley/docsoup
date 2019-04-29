﻿// Copyright (c) Stephen Tetley 2019
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

#load @"ErskineBatteryExtract.fs"
open ErskineBatteryExtract


let localFile (fileName:string) : string = 
    System.IO.Path.Combine (__SOURCE_DIRECTORY__ , "../data", fileName)


let getOk (ans:Result<'a, ErrMsg>) : seq<'a> = 
    match ans with
    | Ok a -> Seq.singleton a
    | Error msg -> printfn "%s" msg; Seq.empty

let processBatch (info:DirectoryInfo) : unit = 
    let outfile = localFile (sprintf "%s erskines.csv" info.Name)
    let okays = 
        System.IO.DirectoryInfo(info.FullName).GetFiles(searchPattern = "*Site Works.docx", searchOption = SearchOption.AllDirectories)
            |> Seq.map (fun file -> 
                            printfn "%s" file.Name ; processSiteWorks file.FullName) 
            |> Seq.collect getOk
    let table = new SiteWorksTable(okays)
    use sw = new StreamWriter(path=outfile, append=false)
    table.Save(writer = sw, separator = ',', quote = '\"')


let main () : unit = 
    let root = @"G:\work\Projects\rtu\Erskines\erskines-incoming"
    Directory.GetDirectories(root) 
        |> Array.iter (fun path -> processBatch (DirectoryInfo(path)))

// Doodle below...


let point : {| X:int; Y:int |} = {| X = 3; Y = 4 |}

let sampleFile = @"G:\work\Projects\rtu\Erskines\erskines-incoming\Batch_02 DU\ALBERT_WTW\ALBERT_WTW Erskine Battery Site Works.docx"

let demo01 () = 
    runDocumentExtractor sampleFile documentInnerText


let dummy1 () = 
    List.find (fun x -> x > 5) [1;2;3;4;5;6;7] 

let dummy2 () = 
    let input = "One Two Three"
    let pattern = "[Tt]wo"
    Regex.IsMatch(input, pattern) 


let demo02 (pattern:string)  = 
    let proc () = 
        body 
            &>> findTable (tableInnerTextMatch pattern)
            &>> tableInnerText
    runDocumentExtractor sampleFile (proc ())
         


let demo03 (number:int)  = 
    let proc () = 
        findM (fun ix -> mreturn (ix = number)) [1;2;3;4;5]
    runDocumentExtractor sampleFile (proc ())

let demo04 ()  = 
    runDocumentExtractor sampleFile (body &>> extractSiteDetails)


let demo05 (pattern:string)  = 
    let proc () = 
        body 
            &>> findTable (tableCell 0 0 &>> cellInnerTextMatch pattern)
            &>> tableInnerText
    runDocumentExtractor sampleFile (proc ())

let demo06 ()  = 
    runDocumentExtractor sampleFile siteWorksExtractor


    