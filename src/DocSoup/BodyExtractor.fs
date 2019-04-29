// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module BodyExtractor = 
    
    open System.Linq

    open DocumentFormat.OpenXml

    open DocSoup

    type BodyExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Body> 

    let (bodyExtractor:BodyExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Body>()

    type BodyExtractor<'a> = ExtractMonad<Wordprocessing.Body,'a> 


    let paragraphs : BodyExtractor<seq<Wordprocessing.Paragraph>> = 
        asks (fun body -> body.OfType<Wordprocessing.Paragraph>())


    let tables : BodyExtractor<seq<Wordprocessing.Table>> = 
        asks (fun body -> body.OfType<Wordprocessing.Table>())

    let table (index:int) : BodyExtractor<Wordprocessing.Table> = 
        bodyExtractor { 
            let! xs = tables
            return! liftOption (Seq.tryItem index xs)
        }

    let findTable (predicate:TableExtractor<bool>) : BodyExtractor<Wordprocessing.Table> = 
        bodyExtractor { 
            let! xs = tables |>> Seq.toList
            return! findM (fun t1 -> (mreturn t1) &>> predicate) xs
        }


    let findTableIndex (predicate:TableExtractor<bool>) : BodyExtractor<int> = 
        bodyExtractor { 
            let! xs = tables |>> Seq.toList
            return! findIndexM (fun t1 -> (mreturn t1) &>> predicate) xs
        }


    let bodyInnerText : BodyExtractor<string> = 
        asks (fun body -> body.InnerText)

