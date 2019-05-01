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
#load @"..\src\DocSoup\Paragraph.fs"
#load @"..\src\DocSoup\Cell.fs"
#load @"..\src\DocSoup\Row.fs"
#load @"..\src\DocSoup\Table.fs"
#load @"..\src\DocSoup\Body.fs"
#load @"..\src\DocSoup\Document.fs"
open DocSoup


#load @"Extractors\Usar\Schema.fs"
#load @"Extractors\Usar\InstallV2.fs"
open Extractors.Usar


let localFile (fileName:string) : string = 
    System.IO.Path.Combine (__SOURCE_DIRECTORY__ , "../data/output", fileName)


let getOk (ans:Result<'a, ErrMsg>) : seq<'a> = 
    match ans with
    | Ok a -> Seq.singleton a
    | Error msg -> printfn "%s" msg; Seq.empty


let getFiles (searchPattern: string) (info:DirectoryInfo) : FileInfo [] =
    System.IO.DirectoryInfo(info.FullName).GetFiles( searchPattern = searchPattern
                                                   , searchOption = SearchOption.AllDirectories)


let processBatch (info:DirectoryInfo) : unit = 
    let outfile = localFile (sprintf "%s usar-installs.csv" info.Name)
    let files1 = getFiles "*Install*.docx" info
    let files2 = getFiles "*Test*.docx" info
    let okays = 
            Array.append files1 files2
            |> Array.toSeq
            |> Seq.map (fun file -> 
                            printfn "%s" file.Name
                            InstallV2.processUsarInstall file.FullName) 
            |> Seq.collect getOk
    let table = new UsarInstallTable(okays)
    use sw = new StreamWriter(path=outfile, append=false)
    table.Save(writer = sw, separator = ',', quote = '\"')

    
let sourceDirectory = @"G:\work\Projects\usar\small-stw\Incoming"

let main () : unit = 
    Directory.GetDirectories(sourceDirectory) 
        |> Array.iter (fun path -> processBatch (DirectoryInfo(path)))
