// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module CellExtract = 
    
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal
    open DocSoup

    type CellExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableCell> 

    let (cellExtractor:CellExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableCell>()

    type CellExtractor<'a> = ExtractMonad<Wordprocessing.TableCell,'a> 

    let cellInnerText : CellExtractor<string> = 
        asks (fun cell -> cell.InnerText)

    /// Get the cell "Paragraphs text" which should preserves newline.
    /// Currently doesn't seem to...
    let cellParagraphsText : CellExtractor<string> = 
        asks (fun cell -> cell.Elements<Wordprocessing.Paragraph>() |> Seq.map (fun text -> text.InnerText)  |> Common.fromLines)
        
    /// This function matches the regex pattern to the 'inner text'
    /// of the cell.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let cellInnerTextMatch (pattern:string) : CellExtractor<bool> = 
        cellExtractor { 
            let! inner = cellInnerText 
            return Regex.IsMatch(inner, pattern)
        }

    let cellParagraphsTextMatch (pattern:string) : CellExtractor<bool> = 
        cellExtractor { 
            let! inner = cellParagraphsText 
            return Regex.IsMatch(inner, pattern)
        }

    let cellIsMatch (pattern:string) : CellExtractor<bool> = 
        cellParagraphsTextMatch pattern

