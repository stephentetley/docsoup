// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"
#r "System.IO.FileSystem.Primitives"

open System.Text.RegularExpressions

#I @"C:\Users\stephen\.nuget\packages\DocumentFormat.OpenXml\2.9.1\lib\netstandard1.3"
#r "DocumentFormat.OpenXml"
#I @"C:\Users\stephen\.nuget\packages\system.io.packaging\4.5.0\lib\netstandard1.3"
#r "System.IO.Packaging"

#load @"..\src\DocSoup\Internal\Common.fs"
#load @"..\src\DocSoup\Internal\OpenXml.fs"
#load @"..\src\DocSoup\ExtractMonad.fs"
#load @"..\src\DocSoup\Combinators.fs"
#load @"..\src\DocSoup\Text.fs"
#load @"..\src\DocSoup\Paragraph.fs"
#load @"..\src\DocSoup\Cell.fs"
#load @"..\src\DocSoup\Row.fs"
#load @"..\src\DocSoup\Table.fs"
#load @"..\src\DocSoup\Body.fs"
#load @"..\src\DocSoup\Document.fs"
open DocSoup

let localFile (fileName:string) : string = 
    System.IO.Path.Combine (__SOURCE_DIRECTORY__ , "../data", fileName)

let testDoc = localFile @"temp-not-for-github.docx"

let demo01 () = 
    Document.runExtractor testDoc Document.innerText



let demo02 () = 
    Document.runExtractor testDoc (Document.body &>> Body.innerText)


let demo03a () : Answer<int> = 
    Document.runExtractor testDoc (Document.body &>> Body.paragraphs &>> asks Seq.length)

let demo03b () : Answer<int> = 
    Document.runExtractor testDoc (Document.body &>> Body.tables &>> asks Seq.length)



let demo04 () : Answer<string> = 
    Document.runExtractor testDoc 
        (Document.body &>> Body.table 0 &>> Table.row 0 &>> Row.innerText)


let demo05a () : Answer<string> = 
    Document.runExtractor testDoc 
        (Document.body &>> Body.table 0 &>> Table.row 14 &>> Row.cell 0 &>> Cell.spacedText)

let demo05b () : Answer<string> = 
    Document.runExtractor testDoc 
        (Document.body &>> Body.table 0 &>> Table.cell (14,0) &>> Cell.spacedText)



let dummy1 () = 
    List.forall (fun x -> x % 2 = 0) [2;4;6]

let dummy1a () = 
    List.pick (fun x -> if x > 7 then Some (x.ToString()) else None) [4;5;6;7;8]


let dummy2 () = 
    let answer = Regex.Match(input = "All the names", pattern = "(?<line1>.*)")
    if answer.Success then
        answer.Groups.["line1"].Value |> Ok
    else
        Error "bah"

let dummy2a () = 
    let answer = Regex.Match(input = "All the names", pattern = ".*")
    if answer.Success then
        answer.Value |> Ok
    else
        Error "bah"

let dummy3 () = 
    let pattern = "(?<one>[Oo]ne).*(?<three>[Tt]hree)"
    let answer = Regex.Match(input = "one two three", pattern = pattern)
    if answer.Success then 
        answer.Groups
            |> Seq.cast<Group> 
            |> Seq.map (fun (g:Group) -> (g.Name, g.Value))
    else    
        failwith "no match"

let dummy3a () = 
    let pattern = "([Oo]ne).*([Tt]hree)"
    let answer = Regex.Match(input = "one two three", pattern = pattern)
    if answer.Success then 
        answer.Groups
            |> Seq.cast<Group> 
            |> Seq.map (fun (g:Group) -> (g.Name, g.Value))
    else    
        failwith "no match"

let dummy4 () = 
    internalRunExtract "one two three" <| Text.matchNamedMatches "(?<one>[Oo]ne).*(?<three>[Tt]hree)"
