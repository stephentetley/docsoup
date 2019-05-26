// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Text2 = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocSoup.Internal.ExtractMonad
    open DocSoup

    type TextExtractorBuilder = ExtractMonadBuilder<string []> 

    let (extractor:TextExtractorBuilder) = new ExtractMonadBuilder<string []>()

    type Extractor<'a> = ExtractMonad<'a, string []> 

    let private contents : Extractor<string []> = asks (fun text -> text)


    let endOfInput : Extractor<unit> = 
        extractor { 
            match! peek "" (fun pos arr -> pos > arr.Length) with
            | true -> return ()
            | false -> do! extractError "end of input"
        }


    /// Caution - use with text extracted with 'spacedText'.
    /// Text extracted with 'innerText' does not preserve line breaks.
    let line : Extractor<string> = 
        consume1 "line read error" (fun ix arr -> arr.[ix])

    let lines (count:int) : Extractor<string []> = 
        consume1 "line read error" (fun ix arr -> arr.[ix .. ix+count])

    let position : Extractor<int> = getPosition ()


    /// Doesn't increase the cursor position.
    let getInput : Extractor<string []> = 
        peek "getInput" (fun ix arr -> arr.[ix ..])
        


    /// Caution - use only on text extracted with 'spacedText'.
    /// Text extracted with 'innerText' does not preserve line breaks.
    let remainingLineCount : Extractor<int> =  
        peek "getInput" (fun ix arr -> arr.Length - ix)


    let satisfy (test:string -> bool) : Extractor<string> = 
        extractor {
            let! ans = line 
            if test ans then 
                return ans 
            else 
                return! extractError "satisfy"
        }
            
    /// Note this works on the current line
    let contains(value:string) : Extractor<string> = 
        satisfy (fun str -> str.Contains(value))

    ///// Caution - use only on text extracted with 'spacedText'.
    ///// Text extracted with 'innerText' does not preserve line breaks.
    //let findLine (predicate:Extractor<bool>) : Extractor<string> = 
    //    extractor { 
    //        let! xs = lines |>> Seq.toList
    //        return! findM (fun line1 -> focus line1 predicate) xs
    //    }

    ///// Caution - use only on text extracted with 'spacedText'.
    ///// Text extracted with 'innerText' does not preserve line breaks.
    //let findLineIndex (predicate:Extractor<bool>) : Extractor<int> = 
    //    extractor { 
    //        let! xs = lines |>> Seq.toList
    //        return! findIndexM (fun cell1 -> focus cell1 predicate) xs
    //    }

    //let subcontents (startIndex: int) (length:int) : Extractor<string> = 
    //    contents |>> fun str -> 
    //        str.Substring(startIndex= startIndex, length = length)

    //let contentsFrom (startIndex: int) : Extractor<string> = 
    //    contents |>> fun str -> 
    //        str.Substring(startIndex= startIndex)

    //let contentsTo (endIndex: int) : Extractor<string> = 
    //    contents |>> fun str -> 
    //        str.Substring(startIndex = 0, length = endIndex)

    //let trim : Extractor<string> = 
    //    contents |>> fun str -> str.Trim()
        

    // ****************************************************
    // Regex matching

    /// Does not consume input.
    let isMatch (pattern:string) : Extractor<bool> = 
        extractor { 
            let! input = lookAhead line 
            let! regexOpts =  getRegexOptions ()
            return Regex.IsMatch( input = input
                                , pattern = pattern
                                , options = regexOpts )
        }

    /// Does not consume input.
    let isNotMatch (pattern:string) : Extractor<bool> = 
        isMatch pattern |>> not

    /// Consumes input
    let regexMatch (pattern:string) : Extractor<RegularExpressions.Match> =
        extractor { 
            let! input = line 
            let! regexOpts =  getRegexOptions ()
            return Regex.Match( input = input
                            , pattern = pattern
                            , options = regexOpts )
        }

    /// Consumes input
    let matchValue (pattern:string) : Extractor<string> =
        regexMatch pattern >>= fun matchObj -> 
        if matchObj.Success then
            mreturn matchObj.Value
        else
            extractError "no match"


    let matchGroups (pattern:string) : Extractor<RegularExpressions.GroupCollection> =
        regexMatch pattern >>= fun matchObj -> 
        if matchObj.Success then
            mreturn matchObj.Groups
        else
            extractError "no match"

    let private isNumber (str:string) : bool = 
        let pattern = "^\d+$"
        Regex.IsMatch(input = str, pattern = pattern)
        
    /// This only returns user named matches, not the 'internal' ones given 
    /// numeric names by .Net's regex library.
    let matchNamedMatches (pattern:string) : Extractor<Map<string, string>> =
        regexMatch pattern >>= fun matchObj -> 
        if matchObj.Success then
            let nameValues = matchObj.Groups |> Seq.cast<Group> 
            let matches = 
                Seq.fold (fun acc (grp:Group) -> 
                            if isNumber grp.Name then acc else Map.add grp.Name grp.Value acc)
                        Map.empty
                        nameValues
            mreturn matches
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

    //let matchStart (pattern:string) : Extractor<int> =
    //    regexMatch pattern >>= fun matchObj -> 
    //    if matchObj.Success then
    //        mreturn matchObj.Index
    //    else
    //        extractError "no match"


    //let matchEnd (pattern:string) : Extractor<int> =
    //    regexMatch pattern >>= fun matchObj -> 
    //    if matchObj.Success then
    //        let start = matchObj.Index
    //        mreturn (start + matchObj.Length)
    //    else
    //        extractError "no match"


    //let leftOfMatch (pattern:string) : Extractor<string> =
    //    matchStart pattern >>= fun ix -> 
    //    contentsTo ix
        

    //let rightOfMatch (pattern:string) : Extractor<string> =
    //    matchEnd pattern >>= fun ix -> 
    //    contentsFrom ix

