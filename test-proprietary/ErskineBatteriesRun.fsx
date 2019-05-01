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
#load @"..\src\DocSoup\Cell.fs"
#load @"..\src\DocSoup\Row.fs"
#load @"..\src\DocSoup\Table.fs"
#load @"..\src\DocSoup\Body.fs"
#load @"..\src\DocSoup\Document.fs"
open DocSoup

#load @"Extractors\ErskineBatteryForm.fs"
open Extractors.ErskineBatteryForm


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
    Document.runExtractor sampleFile Document.innerText


let dummy1 () = 
    List.find (fun x -> x > 5) [1;2;3;4;5;6;7] 

let dummy2 () = 
    let input = "One Two Three"
    let pattern = "[Tt]wo"
    Regex.IsMatch(input, pattern) 


let demo02 (pattern:string)  = 
    let proc () = 
        Document.body 
            &>> Body.findTable (Table.innerTextIsMatch pattern)
            &>> Table.innerText
    Document.runExtractor sampleFile (proc ())
         


let demo03 (number:int)  = 
    let proc () = 
        findM (fun ix -> mreturn (ix = number)) [1;2;3;4;5]
    Document.runExtractor sampleFile (proc ())

let demo04 ()  = 
    Document.runExtractor sampleFile (Document.body &>> extractSiteDetails)


let demo05 (pattern:string)  = 
    let proc () = 
        Document.body 
            &>> Body.findTable (Table.cell (0,0) &>> Cell.innerTextIsMatch pattern)
            &>> Table.innerText
    Document.runExtractor sampleFile (proc ())

let demo06 ()  = 
    Document.runExtractor sampleFile siteWorksExtractor


    