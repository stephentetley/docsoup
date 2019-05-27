// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Paragraphs = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal.ExtractMonad
    open DocSoup


    type ParagraphsExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Paragraph []> 

    let (extractor:ParagraphsExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Paragraph []>()

    type Extractor<'a> = ExtractMonad<'a, Wordprocessing.Paragraph []> 


    let item : Extractor<Wordprocessing.Paragraph> = 
        consume1 "line read error" (fun ix arr -> arr.[ix])

    let paragraphs (count:int) : Extractor<Wordprocessing.Paragraph []> = 
        consume1 "line read error" (fun ix arr -> arr.[ix .. ix+count])

    let position : Extractor<int> = getPosition ()


    /// Doesn't increase the cursor position.
    let getInput : Extractor<Wordprocessing.Paragraph []> = 
        peek "getInput" (fun ix arr -> arr.[ix ..])
    

    let remainingParagraphCount : Extractor<int> =  
        peek "remainingParagraphCount" (fun ix arr -> arr.Length - ix)
