// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Body = 
    
    open System.Linq

    open DocumentFormat.OpenXml

    open DocSoup

    type BodyExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Body> 

    let (extractor:BodyExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Body>()

    type Extractor<'a> = ExtractMonad<Wordprocessing.Body,'a> 


    let paragraphs : Extractor<seq<Wordprocessing.Paragraph>> = 
        asks (fun body -> body.OfType<Wordprocessing.Paragraph>())


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
            return! findM (fun t1 -> (mreturn t1) &>> predicate) xs
        }


    let findTableIndex (predicate:Table.Extractor<bool>) : Extractor<int> = 
        extractor { 
            let! xs = tables |>> Seq.toList
            return! findIndexM (fun t1 -> (mreturn t1) &>> predicate) xs
        }


    let innerText : Extractor<string> = 
        asks (fun body -> body.InnerText)

