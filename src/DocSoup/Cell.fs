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
        
    /// This function matches the regex pattern to the 'inner text'
    /// of the cell.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let innerTextIsMatch (pattern:string) : Extractor<bool> = 
        genRegexIsMatch (fun _ -> innerText) pattern

    /// This function matches the regex pattern to the 'inner text'
    /// of the cell.
    let innerTextIsNotMatch (pattern:string) : Extractor<bool> = 
        innerTextIsMatch pattern |>> not


    let innerTextMatchValue (pattern:string) : Extractor<string> = 
        genRegexMatchValue (fun _ -> innerText) pattern

    let innerTextMatch (pattern:string) : Extractor<RegularExpressions.Match> = 
        genRegexMatch (fun _ -> innerText) pattern

    let innerTextAllMatch (patterns:string []) : Extractor<bool> = 
        genRegexAllMatch (fun _ -> innerText) patterns

    let innerTextAnyMatch (patterns:string []) : Extractor<bool> = 
        genRegexAnyMatch (fun _ -> innerText) patterns

    let spacedTextIsMatch (pattern:string) : Extractor<bool> = 
        genRegexIsMatch (fun _ -> spacedText) pattern

    let spacedTextIsNotMatch (pattern:string) : Extractor<bool> = 
        spacedTextIsMatch pattern |>> not


    let spacedTextMatchValue (pattern:string) : Extractor<string> = 
        genRegexMatchValue (fun _ -> spacedText) pattern

    let spacedTextMatch (pattern:string) : Extractor<RegularExpressions.Match> = 
        genRegexMatch (fun _ -> spacedText) pattern

    let spacedTextAllMatch (patterns:string []) : Extractor<bool> = 
        genRegexAllMatch (fun _ -> spacedText) patterns

    let spacedTextAnyMatch (patterns:string []) : Extractor<bool> = 
        genRegexAnyMatch (fun _ -> spacedText) patterns




