// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"
#r "System.Xml.Linq"
#r "System.IO.FileSystem.Primitives"

open System.IO

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

#load @"Extractors\Mk5ReplacementForm.fs"
open Extractors.Mk5ReplacementForm


let localFile (fileName:string) : string = 
    System.IO.Path.Combine (__SOURCE_DIRECTORY__ , "../data/output", fileName)


let getOk (ans:Result<'a, ErrMsg>) : seq<'a> = 
    match ans with
    | Ok a -> Seq.singleton a
    | Error msg -> printfn "%s" msg; Seq.empty


let main () : unit = 
    let toplevel = @"G:\work\Projects\rtu\mk5-mmims\Incoming"
    let outfile = localFile "mk5-mmim-installs.csv"
    // note some names mispelled as 'Words' not 'Works'...
    let okays = 
        System.IO.DirectoryInfo(toplevel).GetFiles(searchPattern = "*Site Wor?s.docx", searchOption = SearchOption.AllDirectories)
            |> Seq.map (fun file -> 
                            printfn "%s" file.Name ; processMk5Install file.FullName) 
            |> Seq.collect getOk
    let table = new Mk5InstallTable(okays)
    use sw = new StreamWriter(path=outfile, append=false)
    table.Save(writer = sw, separator = ',', quote = '\"')


let sampleFile = @"G:\work\Projects\rtu\mk5-mmims\SAMPLE Upgrade Site Works.docx"

let demo01 () = 
    Document.runExtractor sampleFile (Document.body &>> extractSiteInfo)


let demo02 () = 
    Document.runExtractor sampleFile 
        (Document.body &>> tupleM4 extractSiteInfo 
                                      extractVisitInfo 
                                      extractOutstationInfo
                                      extractAdditionalComments)

