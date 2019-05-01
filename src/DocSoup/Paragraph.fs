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
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let isMatch (pattern:string) : Extractor<bool> = 
        extractor { 
            let! inner = innerText 
            let! regexOpts = getRegexOptions ()
            return Regex.IsMatch( input = inner
                                , pattern = pattern
                                , options = regexOpts )
        }


    let textMatch (pattern:string) : Extractor<RegularExpressions.Match> = 
        extractor { 
            let! inner = innerText 
            let! regexOpts = getRegexOptions ()
            return Regex.Match( input = inner
                              , pattern = pattern
                              , options = regexOpts )
        }


    let matchValue (pattern:string) : Extractor<string> = 
        extractor { 
            let! matchObj = textMatch pattern 
            if matchObj.Success then
                return matchObj.Value
            else
                return! extractError "no match"
        }
