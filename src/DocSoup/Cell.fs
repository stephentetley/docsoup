// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Cell = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal
    open DocSoup

    type CellExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableCell> 

    let (extractor:CellExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableCell>()

    type Extractor<'a> = ExtractMonad<Wordprocessing.TableCell,'a> 

    let innerText : Extractor<string> = 
        asks (fun cell -> cell.InnerText)

    /// Get the cell "Paragraphs text" which should preserves newline.
    let paragraphs : Extractor<seq<Wordprocessing.Paragraph>> = 
        asks (fun cell -> cell.Elements<Wordprocessing.Paragraph>())


    let paragraph (index:int) : Extractor<Wordprocessing.Paragraph> = 
        extractor { 
            let! xs = paragraphs
            return! liftOption (Seq.tryItem index xs)
        }

    /// Get the cell "Paragraphs text" which should preserves newline.
    let spacedText : Extractor<string> = 
        extractor { 
            let! paras = asks (fun cell -> cell.Elements<Wordprocessing.Paragraph>())
            return paras 
                |> Seq.map (fun para1 -> para1.InnerText) 
                |> Common.fromLines
        }
        
    /// This function matches the regex pattern to the 'inner text'
    /// of the cell.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let innerTextIsMatch (pattern:string) : Extractor<bool> = 
        extractor { 
            let! inner = innerText 
            let! regexOpts = getRegexOptions ()
            return Regex.IsMatch( input = inner
                                , pattern = pattern
                                , options = regexOpts )
        }

    let spacedTextIsMatch (pattern:string) : Extractor<bool> = 
        extractor { 
            let! inner = spacedText 
            let! regexOpts = getRegexOptions ()
            return Regex.IsMatch( input = inner
                                , pattern = pattern
                                , options = regexOpts )
        }

    let isMatch (pattern:string) : Extractor<bool> = 
        spacedTextIsMatch pattern


    let spacedTextMatch (pattern:string) : Extractor<RegularExpressions.Match> = 
        extractor { 
            let! inner = spacedText 
            let! regexOpts = getRegexOptions ()
            return Regex.Match( input = inner
                              , pattern = pattern
                              , options = regexOpts )
        }

    let spacedTextMatchValue (pattern:string) : Extractor<string> = 
        extractor { 
            let! matchObj = spacedTextMatch pattern 
            if matchObj.Success then
                return matchObj.Value
            else
                return! extractError "no match"
        }




    let regexMatch (pattern:string) : Extractor<RegularExpressions.Match> = 
        spacedTextMatch pattern

    let regexMatchValue (pattern:string) : Extractor<string> = 
        spacedTextMatchValue pattern