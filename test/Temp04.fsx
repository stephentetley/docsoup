// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


#r "netstandard"

#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"
open Microsoft.Office.Interop

#load @"..\src\DocSoup\Internal\Common.fs"
#load @"..\src\DocSoup\Internal\Region.fs"
open DocSoup.Internal.Common
open DocSoup.Internal

let localFile (fileName:string) : string = 
    System.IO.Path.Combine (__SOURCE_DIRECTORY__ , "../data", fileName)


let temp01 () : bool = 
    Region.contains {Start = 0; End = 100} {Start = 99; End = 100}

let temp02 () : Result<string, string> = 
    let proc (doc:Word.Document) : string = doc.Range() |> cleanRangeText
    primitiveExtract proc (localFile "sample.docx")


let temp03 () : Result<Region.Region option, string> = 
    let proc (doc:Word.Document) : Region.Region option = 
        doc.Range() |> Region.find1 "test" false
    primitiveExtract proc (localFile "sample.docx")



let temp04 () : Result<Region.Region list, string> = 
    let proc (doc:Word.Document) : Region.Region list = 
        doc.Range() |> Region.findMany "one" false 
    primitiveExtract proc (localFile "ones.docx")

