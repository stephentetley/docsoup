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

    type Extractor<'a> = ExtractMonad<'a, Wordprocessing.Document> 

    let runExtractor (filePath:string) (ma:Extractor<'a>) : Result<'a, ErrMsg> = 
        runExtractMonad filePath (fun wpdocument -> wpdocument.MainDocumentPart.Document) ma




    let body : Extractor<Wordprocessing.Body> = 
        asks (fun document -> document.Body)

    let innerText : Extractor<string> = 
        asks (fun document -> document.InnerText)

   