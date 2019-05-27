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
#load @"..\src\DocSoup\Internal\ExtractMonad.fs"
#load @"..\src\DocSoup\Internal\Consume.fs"
#load @"..\src\DocSoup\Combinators.fs"
#load @"..\src\DocSoup\Text.fs"
#load @"..\src\DocSoup\Text2.fs"
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


let demo01 () : Result<_, ErrMsg> =
    Document.runExtractor testDoc 
        (Document.body &>> Body.table 0 &>> Table.spacedText2 &>> Text2.getInput)


let demo02 () : Result<_, ErrMsg> =
    Document.runExtractor testDoc 
        (Document.body &>> Body.table 0 &>> Table.spacedText2 &>> manyTill Text2.getItem (Text2.contains "Site Name"))

let dummy1 () = 
    "Hello".Contains("llo")