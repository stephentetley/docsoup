// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


#r "netstandard"
#r "System.Xml.Linq"
#r "System.IO.FileSystem.Primitives"

open System.IO
open System.Linq
open System.Text.RegularExpressions



#I @"C:\Users\stephen\.nuget\packages\DocumentFormat.OpenXml\2.9.1\lib\netstandard1.3"
#r "DocumentFormat.OpenXml"
#I @"C:\Users\stephen\.nuget\packages\system.io.packaging\4.5.0\lib\netstandard1.3"
#r "System.IO.Packaging"
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing

#load @"..\src\DocSoup\Internal\Common.fs"
#load @"..\src\DocSoup\Internal\OpenXml.fs"
open DocSoup.Internal

let localFile (fileName:string) : string = 
    System.IO.Path.Combine (__SOURCE_DIRECTORY__ , "../data", fileName)


let demo01 () = 
    let document = localFile "ones.docx"
    use wordDoc : WordprocessingDocument =  WordprocessingDocument.Open(document, false)
    use (sr:StreamReader) = new StreamReader(stream = wordDoc.MainDocumentPart.GetStream() )
    let input = sr.ReadToEnd ()
    let regex = new Regex("One")
    regex.Match(input)


let demo02 () = 
    let docPath = localFile "ones.docx"
    OpenXml.primitiveExtract docPath <| fun (wordDoc:WordprocessingDocument) -> 
        let body : Body = wordDoc.MainDocumentPart.Document.Body
        printfn "%s" body.InnerText
        body.OfType<Paragraph>() |> Seq.toList |> List.length



let demo03 () = 
    let docPath = localFile "ones.docx"
    OpenXml.primitiveExtract docPath <| fun (wordDoc:WordprocessingDocument) -> 
        let body : Body = wordDoc.MainDocumentPart.Document.Body
        printfn "%s" body.InnerText
        body.OfType<Paragraph>() |> Seq.iter (fun (para:Paragraph) -> printfn "%s" para.InnerText)




let demo04 () = 
    let docPath = localFile "temp-not-for-github.docx"
    OpenXml.primitiveExtract docPath <| fun (wordDoc:WordprocessingDocument) -> 
        let body : Body = wordDoc.MainDocumentPart.Document.Body
        printfn "%s" body.InnerText
        body |> OpenXml.bodyParagraphs   |> Seq.iter (fun (para:Paragraph) -> printfn "%s" para.InnerText)
        body |> OpenXml.bodyTables       |> Seq.iter (fun (table:Table) -> printfn "%s" table.InnerText)


let cellText (cell:TableCell) : string = 
    cell.Elements<Paragraph>() |> Seq.map (fun text -> text.InnerText)  |> Common.fromLines
    

let printTable (table:Table) : unit = 
    let rows : TableRow seq = table |> OpenXml.tableRows 
    let printRow (row:TableRow) : unit = 
        row |> OpenXml.tableRowCells |> Seq.map cellText |> String.concat "|" |> printfn "%s"
    rows |> Seq.iter printRow
        
let demo05 () = 
    let docPath = localFile "temp-not-for-github.docx"
    OpenXml.primitiveExtract docPath <| fun (wordDoc:WordprocessingDocument) -> 
        let body : Body = wordDoc.MainDocumentPart.Document.Body
        match OpenXml.tryFindTable "Site Details" body with
        | None -> printfn "not found"
        | Some (table:Table) -> printTable table


