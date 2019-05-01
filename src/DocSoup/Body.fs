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

    type Extractor<'a> = ExtractMonad<Wordprocessing.Body,'a> 


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

    /// This function matches the regex pattern to the 'inner text'
    /// of the cell.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let innerTextIsMatch (pattern:string) : Extractor<bool> = 
        genRegexIsMatch (fun _ -> innerText) pattern

    let innerTextMatchValue (pattern:string) : Extractor<string> = 
        genRegexMatchValue (fun _ -> innerText) pattern

    let innerTextMatch (pattern:string) : Extractor<RegularExpressions.Match> = 
        genRegexMatch (fun _ -> innerText) pattern

    let innerTextAllMatch (patterns:string []) : Extractor<bool> = 
        genRegexAllMatch (fun _ -> innerText) patterns

    let innerTextAnyMatch (patterns:string []) : Extractor<bool> = 
        genRegexAnyMatch (fun _ -> innerText) patterns