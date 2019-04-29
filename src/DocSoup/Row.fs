// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Row = 
    
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup

    type RowExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableRow> 

    let (rowExtractor:RowExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableRow>()

    type Extractor<'a> = ExtractMonad<Wordprocessing.TableRow,'a> 

    
    let cells : Extractor<seq<Wordprocessing.TableCell>> = 
        asks (fun row -> row.Elements<Wordprocessing.TableCell>())

    let cell (index:int) : Extractor<Wordprocessing.TableCell> = 
        rowExtractor { 
            let! xs = cells
            return! liftOption (Seq.tryItem index xs)
        }

    let cellCount : Extractor<int> =  
        cells |>> Seq.length

    let firstCell : Extractor<Wordprocessing.TableCell> = 
        cell 0 


    let findCell (predicate:Cell.Extractor<bool>) : Extractor<Wordprocessing.TableCell> = 
        rowExtractor { 
            let! xs = cells |>> Seq.toList
            return! findM (fun t1 -> (mreturn t1) &>> predicate) xs
        }

    let findCellIndex (predicate:Cell.Extractor<bool>) : Extractor<int> = 
        rowExtractor { 
            let! xs = cells |>> Seq.toList
            return! findIndexM (fun t1 -> (mreturn t1) &>> predicate) xs
        }

    let innerText : Extractor<string> = 
        asks (fun row -> row.InnerText)

    /// This function matches the regex pattern to the 'inner text'
    /// of the row.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let innerTextIsMatch (pattern:string) : Extractor<bool> = 
        rowExtractor { 
            let! inner = innerText 
            return Regex.IsMatch(inner, pattern)
        }


    let isMatch (cellPatterns:string []) : Extractor<bool> = 
        rowExtractor { 
            let! arrCells = cells |>> Seq.toArray
            let! pairs = 
                liftAction "zip mismatch" (fun _ -> Array.zip cellPatterns arrCells) |>> Array.toList
            return! forallM (fun (patt,cel) -> local (fun _ -> cel) (Cell.isMatch patt)) pairs
        }

    /// TODO - use regex groups for a function like rowIsMatch that returns matches