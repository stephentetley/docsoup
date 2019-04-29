// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Cell = 
    
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal
    open DocSoup

    type CellExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableCell> 

    let (cellExtractor:CellExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableCell>()

    type Extractor<'a> = ExtractMonad<Wordprocessing.TableCell,'a> 

    let innerText : Extractor<string> = 
        asks (fun cell -> cell.InnerText)

    /// Get the cell "Paragraphs text" which should preserves newline.
    /// Currently doesn't seem to...
    let paragraphsText : Extractor<string> = 
        asks (fun cell -> cell.Elements<Wordprocessing.Paragraph>() |> Seq.map (fun text -> text.InnerText)  |> Common.fromLines)
        
    /// This function matches the regex pattern to the 'inner text'
    /// of the cell.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let innerTextIsMatch (pattern:string) : Extractor<bool> = 
        cellExtractor { 
            let! inner = innerText 
            return Regex.IsMatch(inner, pattern)
        }

    let paragraphsTextIsMatch (pattern:string) : Extractor<bool> = 
        cellExtractor { 
            let! inner = paragraphsText 
            return Regex.IsMatch(inner, pattern)
        }

    let isMatch (pattern:string) : Extractor<bool> = 
        paragraphsTextIsMatch pattern

