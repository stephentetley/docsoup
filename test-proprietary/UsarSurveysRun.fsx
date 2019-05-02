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
#load @"..\src\DocSoup\Text.fs"
#load @"..\src\DocSoup\Paragraph.fs"
#load @"..\src\DocSoup\Cell.fs"
#load @"..\src\DocSoup\Row.fs"
#load @"..\src\DocSoup\Table.fs"
#load @"..\src\DocSoup\Body.fs"
#load @"..\src\DocSoup\Document.fs"
open DocSoup

#load @"Extractors\Usar\Schema.fs"
#load @"Extractors\Usar\SurveyV1.fs"
#load @"Extractors\Usar\SurveyV2.fs"
open Extractors.Usar

let localFile (fileName:string) : string = 
    System.IO.Path.Combine (__SOURCE_DIRECTORY__ , "../data/output", fileName)


let getOk (ans:Result<'a, ErrMsg>) : seq<'a> = 
    match ans with
    | Ok a -> Seq.singleton a
    | Error msg -> printfn "%s" msg; Seq.empty


let processSurvey (path:string) : Answer<UsarSurveyRow> = 
    match SurveyV2.processUsarSurvey path with
    | Ok ans -> Ok ans
    | Error _ -> SurveyV1.processUsarSurvey path



let processBatch (info:DirectoryInfo) : unit = 
    let outfile = localFile (sprintf "%s usar-surveys.csv" info.Name)
    let okays = 
        System.IO.DirectoryInfo(info.FullName).GetFiles(searchPattern = "*Survey*.docx", searchOption = SearchOption.AllDirectories)
            |> Seq.map (fun file -> 
                            printfn "%s" file.Name
                            processSurvey file.FullName) 
            |> Seq.collect getOk
    let table = new UsarSurveyTable(okays)
    use sw = new StreamWriter(path=outfile, append=false)
    table.Save(writer = sw, separator = ',', quote = '\"')

    
let sourceDirectory = @"G:\work\Projects\usar\small-stw\Incoming"

let main () : unit = 
    Directory.GetDirectories(sourceDirectory) 
        |> Array.iter (fun path -> processBatch (DirectoryInfo(path)))

let v1Sample = @"G:\work\Projects\usar\SAMPLE V1 Survey.docx"

let dummy01 () = 
    Document.runExtractor v1Sample 
        (Document.body &>> Body.table 0 &>> Table.firstCell &>> Cell.spacedText)

let dummy02 ()  = 
    Document.runExtractor v1Sample 
            (Document.body &>> 
                tupleM2 SurveyV1.extractSurveyInfo SurveyV1.extractVisitInfo)

let dummy03 () = 
    let ans = Regex.Match(input = "one two three", pattern = "two")
    if ans.Success then
        ans.Index
    else
        failwith "Bad"

let dummy04 () = 
    runExtractMonad v1Sample (fun _ -> "one-two-three")
        <| tupleM2 (Text.leftOfMatch "two") (Text.rightOfMatch "two")

let dummy05 () = 
    runExtractMonad v1Sample (fun _ -> "Site BRUNSWICK")
        <| (Text.rightOfMatch "Site(\s*:?\s*)" &>> Text.trim)

