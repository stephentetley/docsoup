// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module CellExtractor = 
    
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
        
