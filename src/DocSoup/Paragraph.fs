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



    /// This function matches the regex pattern to the 'inner text'
    /// of the cell.
    let textIsMatch (pattern:string) : Extractor<bool> = 
        genRegexIsMatch (fun _ -> innerText) pattern

    /// This function matches the regex pattern to the 'inner text'
    /// of the cell.
    let textMatch (pattern:string) : Extractor<RegularExpressions.Match> = 
        genRegexMatch (fun _ -> innerText) pattern



    let textMatchValue (pattern:string) : Extractor<string> = 
        genRegexMatchValue (fun _ -> innerText) pattern


    let textAllMatch (patterns:string []) : Extractor<bool> = 
        genRegexAllMatch (fun _ -> innerText) patterns


    let textAnyMatch (patterns:string []) : Extractor<bool> = 
        genRegexAnyMatch (fun _ -> innerText) patterns
