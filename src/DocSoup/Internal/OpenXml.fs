// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup.Internal

[<RequireQualifiedAccess>]
module OpenXml = 
    
    open System.Text.RegularExpressions
    open System.Linq
    open DocumentFormat.OpenXml.Packaging
    open DocumentFormat.OpenXml.Wordprocessing
    

    let primitiveExtract (fileName:string) (extract:WordprocessingDocument -> 'a)  : Result<'a, string> =
        if System.IO.File.Exists (fileName) then
            use wordDoc : WordprocessingDocument =  WordprocessingDocument.Open(path = fileName, isEditable = false)
            let ans = extract wordDoc
            Ok ans
        else 
            Error <| sprintf "Cannot find file: %s" fileName


    let documentBody (wordDoc:WordprocessingDocument) : Body = 
        wordDoc.MainDocumentPart.Document.Body



    let inline bodyTables(body:Body) : Table seq = 
        body.OfType<Table>()

    let inline bodyTable(body:Body) (index:int) : Table = 
        bodyTables body |> Seq.item index

    let tableRows(table:Table) : TableRow seq = 
        table.Elements<TableRow>() 

    let tableRowCells (tr:TableRow) : TableCell seq = 
        tr.Elements<TableCell>()

    let inline bodyParagraphs(body:Body) : Paragraph seq = 
        body.OfType<Paragraph>()

    let bodyParagraph(body:Body) (index:int) : Paragraph = 
        bodyParagraphs body |> Seq.item index


    
    let tableMatch (pattern:string) (table:Table)  : Match = 
        Regex.Match(table.InnerText, pattern)

    let paragraphMatch (pattern:string) (paragraph:Paragraph)  : Match = 
        Regex.Match(paragraph.InnerText, pattern)


    // Find the first table whose innerText matches the pattern
    let tryFindTable (pattern:string) (body:Body) : Table option = 
        bodyTables body |> Seq.tryFind (fun table -> (tableMatch pattern table).Success)

    let findTables (pattern:string) (body:Body) : Table seq = 
        bodyTables body |> Seq.filter (fun table -> (tableMatch pattern table).Success)


    // Find the first table whose innerText matches the pattern
    let tryFindParagraph (pattern:string) (body:Body) : Paragraph option = 
        bodyParagraphs body |> Seq.tryFind (fun paragraph -> (paragraphMatch pattern paragraph).Success)

        
    let findParagraph (pattern:string) (body:Body) : Paragraph seq = 
        bodyParagraphs body |> Seq.filter (fun paragraph -> (paragraphMatch pattern paragraph).Success)