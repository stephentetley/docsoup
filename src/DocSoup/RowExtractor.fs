// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module RowExtractor = 
    
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup

    type RowExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableRow> 

    let (rowExtractor:RowExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableRow>()

    type RowExtractor<'a> = ExtractMonad<Wordprocessing.TableRow,'a> 

    
    let cells : RowExtractor<seq<Wordprocessing.TableCell>> = 
        asks (fun row -> row.Elements<Wordprocessing.TableCell>())

    let cell (index:int) : RowExtractor<Wordprocessing.TableCell> = 
        rowExtractor { 
            let! xs = cells
            return! liftOption (Seq.tryItem index xs)
        }

    let findCell (predicate:CellExtractor<bool>) : RowExtractor<Wordprocessing.TableCell> = 
        rowExtractor { 
            let! xs = cells |>> Seq.toList
            return! findM (fun t1 -> (mreturn t1) &>> predicate) xs
        }

    let findCellIndex (predicate:CellExtractor<bool>) : RowExtractor<int> = 
        rowExtractor { 
            let! xs = cells |>> Seq.toList
            return! findIndexM (fun t1 -> (mreturn t1) &>> predicate) xs
        }

    let rowInnerText : RowExtractor<string> = 
        asks (fun row -> row.InnerText)

    /// This function matches the regex pattern to the 'inner text'
    /// of the row.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let rowInnerTextMatch (pattern:string) : RowExtractor<bool> = 
        rowExtractor { 
            let! inner = rowInnerText 
            return Regex.IsMatch(inner, pattern)
        }
