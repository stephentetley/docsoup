﻿// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Row = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal
    open DocSoup.Internal.ExtractMonad
    open DocSoup

    type RowExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableRow> 

    let (extractor:RowExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableRow>()

    type Extractor<'a> = ExtractMonad<'a, Wordprocessing.TableRow> 

    
    let cells : Extractor<seq<Wordprocessing.TableCell>> = 
        asks (fun row -> row.Elements<Wordprocessing.TableCell>())

    let cell (index:int) : Extractor<Wordprocessing.TableCell> = 
        extractor { 
            let! xs = cells
            return! liftOption (Seq.tryItem index xs)
        }

    let cellCount : Extractor<int> =  
        cells |>> Seq.length

    let firstCell : Extractor<Wordprocessing.TableCell> = 
        cell 0 


    let findCell (predicate:Cell.Extractor<bool>) : Extractor<Wordprocessing.TableCell> = 
        extractor { 
            let! xs = cells |>> Seq.toList
            return! findM (fun cell1 -> focus cell1 predicate) xs
        }

    let findCellIndex (predicate:Cell.Extractor<bool>) : Extractor<int> = 
        extractor { 
            let! xs = cells |>> Seq.toList
            return! findIndexM (fun cell1 -> focus cell1 predicate) xs
        }

    let innerText : Extractor<string> = 
        asks (fun row -> row.InnerText)


    /// Get the cell "Spaced text" which should preserves newline.
    let cellsSpacedText : Extractor<string []> = 
        extractor { 
            let! cellList = cells |>> Seq.toList
            let! texts = mapM (fun cell1 -> focus cell1 Cell.spacedText) cellList
            return texts |> List.toArray
        }

    let spacedText : Extractor<string> = 
        cellsSpacedText |>> Common.fromLines


    let cellsMatch (cellPatterns:string []) : Extractor<bool> = 
        extractor { 
            let! arrCells = cells |>> Seq.toArray
            let! pairs = 
                liftOperation "zip mismatch" (fun _ -> Array.zip cellPatterns arrCells) |>> Array.toList
            return! forallM (fun (patt,cell1) -> focus cell1 (Cell.spacedText &>> Text.isMatch patt)) pairs
        }

    // TODO - use regex groups for a function like rowIsMatch that returns matches

    /// Use with caution - a `RegularExpressions.Match` has not necessarily 
    /// matched the input string. The `.Success` property may be false.
    let cellsRegexMatch (cellPatterns:string []) : Extractor<RegularExpressions.Match []> = 
        extractor { 
            let! arrCells = cells |>> Seq.toArray
            let! pairs = 
                liftOperation "zip mismatch" (fun _ -> Array.zip cellPatterns arrCells) |>> Array.toList
            return! mapM (fun (patt,cell1) -> focus cell1 (Cell.spacedText &>> Text.regexMatch patt)) pairs |>> List.toArray
        }


    /// Prefer this to `regexMatch` if you are expecting an array 
    /// of successful matches and you don't need to inspect the result
    /// (e.g. for a match group).
    let cellsMatchValues (cellPatterns:string []) : Extractor<string []> = 
        extractor { 
            let! arrCells = cells |>> Seq.toArray
            let! pairs = 
                liftOperation "zip mismatch" (fun _ -> Array.zip cellPatterns arrCells) |>> Array.toList
            return! mapM (fun (patt,cell1) -> 
                            focus cell1 (Cell.spacedText &>> Text.matchValue patt)) pairs |>> List.toArray
        }

    /// Parse a two column row with "name" in the first cell and 
    /// "value" in the second cell.
    let nameValue2Row (namePattern:string) : Extractor<string> = 
        extractor { 
            let! arr = cellsMatchValues [| namePattern; ".*" |]
            let! ans = liftOperation "nameValue1Row - bad index" (fun _ -> arr.[1])
            return ans
        }

    /// Parse a three column row with "name" in the first cell and 
    /// "value1" and "value2" in the second and third cells.
    let nameValue3Row (namePattern:string) : Extractor<string * string> = 
        extractor { 
            let! arr = cellsMatchValues [| namePattern; ".*"; ".*" |]
            let! ans1 = liftOperation "nameValue2Row - bad index" (fun _ -> arr.[1])
            let! ans2 = liftOperation "nameValue2Row - bad index" (fun _ -> arr.[2])
            return (ans1, ans2)
        }

    /// Parse a four column row with "name" in the first cell and 
    /// "value1", "value2" and "value3" in the second, third and fourth cells.
    let nameValue4Row (namePattern:string) : Extractor<string * string * string> = 
        extractor { 
            let! arr = cellsMatchValues [| namePattern; ".*"; ".*" ; ".*" |]
            let! ans1 = liftOperation "nameValue3Row - bad index" (fun _ -> arr.[1])
            let! ans2 = liftOperation "nameValue3Row - bad index" (fun _ -> arr.[2])
            let! ans3 = liftOperation "nameValue3Row - bad index" (fun _ -> arr.[3])
            return (ans1, ans2, ans3)
        }

    /// TODO - nameValuesRow (namePattern:string) : Extractor<string []> 