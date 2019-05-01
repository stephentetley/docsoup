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
    let paragraphsText : Extractor<string> = 
        extractor { 
            let! paras = asks (fun cell -> cell.Elements<Wordprocessing.Paragraph>())
            return paras |> Seq.map (fun para -> para.InnerText) |> Common.fromLines
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

    let paragraphsTextIsMatch (pattern:string) : Extractor<bool> = 
        extractor { 
            let! inner = paragraphsText 
            let! regexOpts = getRegexOptions ()
            return Regex.IsMatch( input = inner
                                , pattern = pattern
                                , options = regexOpts )
        }

    let isMatch (pattern:string) : Extractor<bool> = 
        paragraphsTextIsMatch pattern


    let paragraphsTextMatch (pattern:string) : Extractor<RegularExpressions.Match> = 
        extractor { 
            let! inner = paragraphsText 
            let! regexOpts = getRegexOptions ()
            return Regex.Match( input = inner
                              , pattern = pattern
                              , options = regexOpts )
        }

    let paragraphsTextMatchValue (pattern:string) : Extractor<string> = 
        extractor { 
            let! matchObj = paragraphsTextMatch pattern 
            if matchObj.Success then
                return matchObj.Value
            else
                return! extractError "no match"
        }




    let regexMatch (pattern:string) : Extractor<RegularExpressions.Match> = 
        paragraphsTextMatch pattern

    let regexMatchValue (pattern:string) : Extractor<string> = 
        paragraphsTextMatchValue pattern