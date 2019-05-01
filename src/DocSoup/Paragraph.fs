// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Paragraph = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal
    open DocSoup

    type ParagraphExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Paragraph> 

    let (extractor:ParagraphExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Paragraph>()

    type Extractor<'a> = ExtractMonad<Wordprocessing.Paragraph,'a> 

    let innerText : Extractor<string> = 
        asks (fun paragraph -> paragraph.InnerText)


