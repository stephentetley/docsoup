// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Table = 
    
    open DocumentFormat.OpenXml

    
    open DocSoup.Internal
    open DocSoup.Internal.ExtractMonad
    open DocSoup

    type TableExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Table> 

    let (extractor:TableExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Table>()

    type Extractor<'a> = ExtractMonad<'a, Wordprocessing.Table> 



    let rows : Extractor<seq<Wordprocessing.TableRow>> = 
        asks (fun table -> table.Elements<Wordprocessing.TableRow>())

    let row (index:int) : Extractor<Wordprocessing.TableRow> = 
        extractor { 
            let! xs = rows
            return! liftOption (Seq.tryItem index xs)
        }

    let rowCount : Extractor<int> = rows |>> Seq.length

    /// Get the structure of a table which is an array 
    /// of column counts for each row.
    let structure : Extractor<int []> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            let! counts = mapM (fun row1 -> focus row1 Row.cellCount) xs
            return counts |> List.toArray
        }


    let cell (rowIndex:int, columnIndex:int) : Extractor<Wordprocessing.TableCell> = 
        extractor { 
            let! xs = row rowIndex 
                        |>> fun row1 -> row1.Elements<Wordprocessing.TableCell>()
            return! liftOption (Seq.tryItem columnIndex xs)
        }

    let firstRow : Extractor<Wordprocessing.TableRow> = row 0  

    let firstCell : Extractor<Wordprocessing.TableCell> = cell (0,0)


    let findRow (predicate:Row.Extractor<bool>) : Extractor<Wordprocessing.TableRow> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            return! findM (fun row1 -> focus row1 predicate) xs
        }

    let findRowIndex (predicate:Row.Extractor<bool>) : Extractor<int> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            return! findIndexM (fun row1 -> focus row1 predicate) xs
        }


    let innerText : Extractor<string> = 
        asks (fun table -> table.InnerText)

    /// Get the row "Paragraphs text" which should preserves newline.
    let rowsSpacedText : Extractor<string []> = 
        extractor { 
            let! rowList = rows |>> Seq.toList
            let! texts = mapM (fun row1 -> focus row1 Row.spacedText) rowList
            return texts |> List.toArray
        }

    let spacedText : Extractor<string> = 
        rowsSpacedText |>> Common.fromLines



    /// Find the string value in a two cell row.
    /// Cell 1 is considered the "name" field.
    /// Cell 2 is "value".
    let findNameValue2Row (namePattern:string) : Extractor<string> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            return! pickM (fun row1 -> focus row1 <| Row.nameValue2Row namePattern) xs
        }

    /// Find the string values in a three cell row.
    /// Cell 1 is considered the "name" field.
    /// Cell 2 is "value 1".
    /// Cell 3 is "value 2".
    let findNameValue3Row (namePattern:string) : Extractor<string * string> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            return! pickM (fun row1 -> focus row1 <| Row.nameValue3Row namePattern) xs
        }
        

    /// Find the string values in a four cell row.
    /// Cell 1 is considered the "name" field.
    /// Cell 2 is "value 1".
    /// Cell 3 is "value 2".
    /// Cell 4 is "value 3".
    let findNameValue4Row (namePattern:string) : Extractor<string * string * string> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            return! pickM (fun row1 -> focus row1 <| Row.nameValue4Row namePattern) xs
        }