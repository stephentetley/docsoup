// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"
#r "System.Xml.Linq"

open System.IO


#I @"C:\Users\stephen\.nuget\packages\system.io.filesystem.primitives\4.3.0\lib\netstandard1.3"
#r "System.IO.FileSystem.Primitives"

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
#load @"..\src\DocSoup\DocumentExtractor.fs"
#load @"..\src\DocSoup\BodyExtractor.fs"
#load @"..\src\DocSoup\TableExtractor.fs"
#load @"..\src\DocSoup\RowExtractor.fs"
#load @"..\src\DocSoup\CellExtractor.fs"
open DocSoup

#load @"SurveyExtractor.fs"
open SurveyExtractor


let localFile (fileName:string) : string = 
    System.IO.Path.Combine (__SOURCE_DIRECTORY__ , "../data", fileName)


let getOk (ans:Result<'a, ErrMsg>) : seq<'a> = 
    match ans with
    | Ok a -> Seq.singleton a
    | Error msg -> printfn "%s" msg; Seq.empty

let main () : unit = 
    let root = @"G:\work\Projects\uqpb\commission-form-samples"
    let outfile = localFile "surveys.csv"
    let files = System.IO.DirectoryInfo(root).GetFiles(searchPattern = "*.docx")
        
    let okays = files |> Seq.map (fun file -> processSurvey file.FullName) |> Seq.collect getOk
    let table = new SurveyTable(okays)
    use sw = new StreamWriter(path=outfile, append=false)
    table.Save(writer = sw, separator = ',', quote = '\"')
