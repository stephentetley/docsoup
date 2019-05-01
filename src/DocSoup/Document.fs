// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Document = 
    
    open System.Text
    open DocumentFormat.OpenXml

    open DocSoup

    type DocumentExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Document> 

    let (extractor:DocumentExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Document>()

    type Extractor<'a> = ExtractMonad<Wordprocessing.Document,'a> 

    let runExtractor (filePath:string) (ma:Extractor<'a>) : Answer<'a> = 
        runExtractMonad filePath (fun wpdocument -> wpdocument.MainDocumentPart.Document) ma



    let body : Extractor<Wordprocessing.Body> = 
        asks (fun document -> document.Body)

    let innerText : Extractor<string> = 
        asks (fun document -> document.InnerText)

    /// This function matches the regex pattern to the 'inner text'
    /// of the cell.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let innerTextIsMatch (pattern:string) : Extractor<bool> = 
        genRegexIsMatch (fun _ -> innerText) pattern

    let innerTextIsNotMatch (pattern:string) : Extractor<bool> = 
        innerTextIsMatch pattern |>> not

    let innerTextMatchValue (pattern:string) : Extractor<string> = 
        genRegexMatchValue (fun _ -> innerText) pattern

    let innerTextMatch (pattern:string) : Extractor<RegularExpressions.Match> = 
        genRegexMatch (fun _ -> innerText) pattern

    let innerTextAllMatch (patterns:string []) : Extractor<bool> = 
        genRegexAllMatch (fun _ -> innerText) patterns

    let innerTextAnyMatch (patterns:string []) : Extractor<bool> = 
        genRegexAnyMatch (fun _ -> innerText) patterns