// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"


#I @"C:\Users\stephen\.nuget\packages\system.io.filesystem.primitives\4.3.0\lib\netstandard1.3"
#r "System.IO.FileSystem.Primitives"

#I @"C:\Users\stephen\.nuget\packages\DocumentFormat.OpenXml\2.9.1\lib\netstandard1.3"
#r "DocumentFormat.OpenXml"
#I @"C:\Users\stephen\.nuget\packages\system.io.packaging\4.5.0\lib\netstandard1.3"
#r "System.IO.Packaging"
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing


#load @"..\src\DocSoup\Internal\Common.fs"
#load @"..\src\DocSoup\Internal\OpenXml.fs"
#load @"..\src\DocSoup\ExtractMonad.fs"
#load @"..\src\DocSoup\DocumentExtractor.fs"
#load @"..\src\DocSoup\BodyExtractor.fs"
#load @"..\src\DocSoup\TableExtractor.fs"
#load @"..\src\DocSoup\RowExtractor.fs"
#load @"..\src\DocSoup\CellExtractor.fs"
open DocSoup


let testDoc = @"E:\coding\fsharp\docsoup\data\temp-not-for-github.docx"

let demo01 () = 
    runDocumentExtractor testDoc documentInnerText



let demo02 () = 
    runDocumentExtractor testDoc (body &>> bodyInnerText)


let demo03a () : Answer<int> = 
    runDocumentExtractor testDoc (body &>> paragraphs &>> asks Seq.length)

let demo03b () : Answer<int> = 
    runDocumentExtractor testDoc (body &>> tables &>> asks Seq.length)



let demo04 () : Answer<string> = 
    runDocumentExtractor testDoc (body &>> table 0 &>> row 0 &>> rowInnerText)


let demo05a () : Answer<string> = 
    runDocumentExtractor testDoc (body &>> table 0 &>> row 14 &>> cell 0 &>> cellParagraphsText)

let demo05b () : Answer<string> = 
    runDocumentExtractor testDoc (body &>> table 0 &>> tableCell 14 0 &>> cellParagraphsText)

