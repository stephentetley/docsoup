// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module TableExtract = 
    
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup

    type TableExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Table> 

    let (tableExtractor:TableExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Table>()

    type TableExtractor<'a> = ExtractMonad<Wordprocessing.Table,'a> 



    let rows : TableExtractor<seq<Wordprocessing.TableRow>> = 
        asks (fun table -> table.Elements<Wordprocessing.TableRow>())

    let row (index:int) : TableExtractor<Wordprocessing.TableRow> = 
        tableExtractor { 
            let! xs = rows
            return! liftOption (Seq.tryItem index xs)
        }

    let rowCount : TableExtractor<int> = rows |>> Seq.length


    let tableCell (rowIndex:int) (columnIndex:int) : TableExtractor<Wordprocessing.TableCell> = 
        tableExtractor { 
            let! xs = row rowIndex |>> fun r1 -> r1.Elements<Wordprocessing.TableCell>()
            return! liftOption (Seq.tryItem columnIndex xs)
        }


    let findRow (predicate:RowExtractor<bool>) : TableExtractor<Wordprocessing.TableRow> = 
        tableExtractor { 
            let! xs = rows |>> Seq.toList
            return! findM (fun t1 -> (mreturn t1) &>> predicate) xs
        }

    let findRowIndex (predicate:RowExtractor<bool>) : TableExtractor<int> = 
        tableExtractor { 
            let! xs = rows |>> Seq.toList
            return! findIndexM (fun t1 -> (mreturn t1) &>> predicate) xs
        }


    let tableInnerText : TableExtractor<string> = 
        asks (fun table -> table.InnerText)

    /// This function matches the regex pattern to the 'inner text'
    /// of the table.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let tableInnerTextMatch (pattern:string) : TableExtractor<bool> = 
        tableExtractor { 
            let! inner = tableInnerText 
            return Regex.IsMatch(inner, pattern)
        }