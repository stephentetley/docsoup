// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Document = 
    
    open DocumentFormat.OpenXml

    open DocSoup

    type DocumentExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Document> 

    let (extractor:DocumentExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Document>()

    type Extractor<'a> = ExtractMonad<Wordprocessing.Document,'a> 

    let runExtractor (filePath:string) (ma:Extractor<'a>) : Answer<'a> = 
        runExtractMonad filePath (fun wpdocument -> wpdocument.MainDocumentPart.Document) ma

    let innerText : Extractor<string> = 
        asks (fun document -> document.InnerText)

    let body : Extractor<Wordprocessing.Body> = 
        asks (fun document -> document.Body)

