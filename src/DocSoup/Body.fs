// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Body = 
   
    open System.Text
    open System.Linq

    open DocumentFormat.OpenXml

    open DocSoup

    type BodyExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Body> 

    let (extractor:BodyExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Body>()

    type Extractor<'a> = ExtractMonad<'a, Wordprocessing.Body> 


    let paragraphs : Extractor<seq<Wordprocessing.Paragraph>> = 
        asks (fun body -> body.OfType<Wordprocessing.Paragraph>())


    let paragraph (index:int) : Extractor<Wordprocessing.Paragraph> = 
        extractor { 
            let! xs = paragraphs
            return! liftOption (Seq.tryItem index xs)
        }


    let paragraphCount : Extractor<int> = paragraphs |>> Seq.length

    let firstParagraph : Extractor<Wordprocessing.Paragraph> = paragraph 0 


    let findParagraph (predicate:Paragraph.Extractor<bool>) : Extractor<Wordprocessing.Paragraph> = 
        extractor { 
            let! xs = paragraphs |>> Seq.toList
            return! findM (fun para1 -> focus para1 predicate) xs
        }

    let findParagraphIndex (predicate:Paragraph.Extractor<bool>) : Extractor<int> = 
        extractor { 
            let! xs = paragraphs |>> Seq.toList
            return! findIndexM (fun para1 -> focus para1 predicate) xs
        }


    let tables : Extractor<seq<Wordprocessing.Table>> = 
        asks (fun body -> body.OfType<Wordprocessing.Table>())

    let table (index:int) : Extractor<Wordprocessing.Table> = 
        extractor { 
            let! xs = tables
            return! liftOption (Seq.tryItem index xs)
        }

    let findTable (predicate:Table.Extractor<bool>) : Extractor<Wordprocessing.Table> = 
        extractor { 
            let! xs = tables |>> Seq.toList
            return! findM (fun table1 -> focus table1 predicate) xs
        }


    let findTableIndex (predicate:Table.Extractor<bool>) : Extractor<int> = 
        extractor { 
            let! xs = tables |>> Seq.toList
            return! findIndexM (fun table1 -> focus table1 predicate) xs
        }


    let innerText : Extractor<string> = 
        asks (fun body -> body.InnerText)
