// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Cell = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal
    open DocSoup.Internal.ExtractMonad
    open DocSoup

    type CellExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableCell> 

    let (extractor:CellExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableCell>()

    type Extractor<'a> = ExtractMonad<'a, Wordprocessing.TableCell> 


    /// Get Paragraphs in the cell.
    let paragraphs : Extractor<seq<Wordprocessing.Paragraph>> = 
        asks (fun cell -> cell.Elements<Wordprocessing.Paragraph>())


    let paragraph (index:int) : Extractor<Wordprocessing.Paragraph> = 
        extractor { 
            let! xs = paragraphs
            return! liftOption (Seq.tryItem index xs)
        }

    let paragraphCount : Extractor<int> = paragraphs |>> Seq.length

    let firstParagraph : Extractor<Wordprocessing.Paragraph> = paragraph 0 


    let findParagraph (predicate:Paragraph.Extractor<bool>) : Extractor<Wordprocessing.Paragraph> = 
        extractor { 
            let! xs = paragraphs |>> Seq.toList
            return! findM (fun para1 -> focus para1 predicate) xs
        }

    let findParagraphIndex (predicate:Paragraph.Extractor<bool>) : Extractor<int> = 
        extractor { 
            let! xs = paragraphs |>> Seq.toList
            return! findIndexM (fun para1 -> focus para1 predicate) xs
        }

    // ****************************************************
    // Get the text

    let innerText : Extractor<string> = 
        asks (fun cell -> cell.InnerText)


    /// Get the cell "Paragraphs text" which should preserves newline.
    let spacedText : Extractor<string> = 
        extractor { 
            let! paras = asks (fun cell -> cell.Elements<Wordprocessing.Paragraph>())
            return paras 
                |> Seq.map (fun para1 -> para1.InnerText) 
                |> Common.fromLines
        }
        
    /// Get the cell "Paragraphs text lines"
    let spacedText2 : Extractor<string []> = 
        extractor { 
            let! paras = asks (fun cell -> cell.Elements<Wordprocessing.Paragraph>())
            return paras 
                |> Seq.map (fun para1 -> para1.InnerText) 
                |> Seq.toArray
        }



