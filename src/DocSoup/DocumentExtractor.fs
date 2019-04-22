// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module DocumentExtractor = 
    
    open DocumentFormat.OpenXml

    open DocSoup

    type DocumentExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Document> 

    let (documentExtractor:DocumentExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Document>()

    type DocumentExtractor<'a> = ExtractMonad<Wordprocessing.Document,'a> 

    let runDocumentExtractor (filePath:string) (ma:DocumentExtractor<'a>) : Answer<'a> = 
        runExtractMonad filePath (fun wpdocument -> wpdocument.MainDocumentPart.Document) ma

    let documentInnerText : DocumentExtractor<string> = 
        asks (fun document -> document.InnerText)

    let body : DocumentExtractor<Wordprocessing.Body> = 
        asks (fun document -> document.Body)

