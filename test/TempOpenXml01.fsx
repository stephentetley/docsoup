// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


#r "netstandard"
#r "System.Xml.Linq"


open System.IO
open System.Text.RegularExpressions
open System.Linq

#I @"C:\Users\stephen\.nuget\packages\DocumentFormat.OpenXml\2.9.1\lib\netstandard1.3"
#r "DocumentFormat.OpenXml"
#I @"C:\Users\stephen\.nuget\packages\system.io.packaging\4.5.0\lib\netstandard1.3"
#r "System.IO.Packaging"
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing

let localFile (fileName:string) : string = 
    System.IO.Path.Combine (__SOURCE_DIRECTORY__ , "../data", fileName)


let demo01 () = 
    let document = localFile "ones.docx"
    use wordDoc : WordprocessingDocument =  WordprocessingDocument.Open(document, false)
    use (sr:StreamReader) = new StreamReader(wordDoc.MainDocumentPart.GetStream() )
    let input = sr.ReadToEnd ()
    let regex = new Regex("One")
    regex.Match(input)


let demo02 () = 
    let document = localFile "ones.docx"
    use wordDoc : WordprocessingDocument =  WordprocessingDocument.Open(document, false)
    let body : Body = wordDoc.MainDocumentPart.Document.Body
    printfn "%s" body.InnerText
    body.OfType<Paragraph>() |> Seq.toList |> List.length



