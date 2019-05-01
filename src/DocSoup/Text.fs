// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Text = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal
    open DocSoup

    type TextExtractorBuilder = ExtractMonadBuilder<string> 

    let (extractor:TextExtractorBuilder) = new ExtractMonadBuilder<string>()

    type Extractor<'a> = ExtractMonad<string,'a> 

    let contents : Extractor<string> = 
        asks (fun text -> text)

    let isMatch (pattern:string) : Extractor<bool> = 
        genRegexIsMatch (fun _ -> contents) pattern

    let isNotMatch (pattern:string) : Extractor<bool> = 
        isMatch pattern |>> not


    let matchValue (pattern:string) : Extractor<string> = 
        genRegexMatchValue (fun _ -> contents) pattern


    let allMatch (patterns:string []) : Extractor<bool> = 
        genRegexAllMatch (fun _ -> contents) patterns


    let anyMatch (patterns:string []) : Extractor<bool> = 
        genRegexAnyMatch (fun _ -> contents) patterns

