// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Paragraph = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal.ExtractMonad
    open DocSoup


    type ParagraphExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Paragraph> 

    let (extractor:ParagraphExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Paragraph>()

    type Extractor<'a> = ExtractMonad<'a, Wordprocessing.Paragraph> 

    let content : Extractor<Wordprocessing.Paragraph> = 
        asks id


    let satisfy (test:Wordprocessing.Paragraph -> bool) : Extractor<Wordprocessing.Paragraph> = 
        extractor {
            let! ans = content 
            if test ans then 
                return ans 
            else 
                return! extractError "satisfy"
        }


    let innerText : Extractor<string> = 
        asks (fun paragraph -> paragraph.InnerText)


