// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Text = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocSoup.Internal
    open DocSoup

    type TextExtractorBuilder = ExtractMonadBuilder<string> 

    let (extractor:TextExtractorBuilder) = new ExtractMonadBuilder<string>()

    type Extractor<'a> = ExtractMonad<string,'a> 

    let contents : Extractor<string> = asks (fun text -> text)


    let lines () : Extractor<string []> = contents |>> Common.toLines

    let subcontents (startIndex: int) (length:int) : Extractor<string> = 
        contents |>> fun str -> 
            str.Substring(startIndex= startIndex, length = length)

    let contentsFrom (startIndex: int) : Extractor<string> = 
        contents |>> fun str -> 
            str.Substring(startIndex= startIndex)

    let contentsTo (endIndex: int) : Extractor<string> = 
        contents |>> fun str -> 
            str.Substring(startIndex = 0, length = endIndex)


    // ****************************************************
    // Regex matching

    let isMatch (pattern:string) : Extractor<bool> = 
        extractor { 
            let! input = contents 
            let! regexOpts =  getRegexOptions ()
            return Regex.IsMatch( input = input
                                , pattern = pattern
                                , options = regexOpts )
        }


    let isNotMatch (pattern:string) : Extractor<bool> = 
        isMatch pattern |>> not


    let regexMatch (pattern:string) : Extractor<RegularExpressions.Match> =
        extractor { 
            let! input = contents 
            let! regexOpts =  getRegexOptions ()
            return Regex.Match( input = input
                            , pattern = pattern
                            , options = regexOpts )
        }

    let matchValue (pattern:string) : Extractor<string> =
        regexMatch pattern >>= fun matchObj -> 
        if matchObj.Success then
            mreturn matchObj.Value
        else
            extractError "no match"


    let anyMatch (patterns:string []) : Extractor<bool> = 
        let (predicates : Extractor<bool> list) = 
            patterns |> Array.toList |> List.map isMatch
        anyM predicates


    let allMatch (patterns:string []) : Extractor<bool> = 
        let (predicates : Extractor<bool> list) = 
            patterns |> Array.toList |> List.map isMatch
        allM predicates

    let matchStart (pattern:string) : Extractor<int> =
        regexMatch pattern >>= fun matchObj -> 
        if matchObj.Success then
            mreturn matchObj.Index
        else
            extractError "no match"


    let matchEnd (pattern:string) : Extractor<int> =
        regexMatch pattern >>= fun matchObj -> 
        if matchObj.Success then
            let start = matchObj.Index
            mreturn (start + matchObj.Length)
        else
            extractError "no match"


    let leftOfMatch (pattern:string) : Extractor<string> =
        matchStart pattern >>= fun ix -> 
        contentsTo ix
        

    let rightOfMatch (pattern:string) : Extractor<string> =
        matchEnd pattern >>= fun ix -> 
        contentsFrom ix

