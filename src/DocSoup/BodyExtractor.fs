// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module BodyExtractor = 
    
    open DocumentFormat.OpenXml

    open DocSoup

    type BodyExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Body> 

    let (bodyExtractor:BodyExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Body>()

    type BodyExtractor<'a> = ExtractMonad<Wordprocessing.Body,'a> 

    let bodyInnerText : BodyExtractor<string> = 
        asks (fun body -> body.InnerText)

