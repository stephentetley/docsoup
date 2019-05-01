// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Row = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup
    open DocSoup.Internal

    type RowExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableRow> 

    let (extractor:RowExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableRow>()

    type Extractor<'a> = ExtractMonad<Wordprocessing.TableRow,'a> 

    
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


    /// Get the cell "Paragraphs text" which should preserves newline.
    let cellsParagraphsText : Extractor<string []> = 
        extractor { 
            let! cellList = cells |>> Seq.toList
            let! texts = mapM (fun cell1 -> focus cell1 Cell.paragraphsText) cellList
            return texts |> List.toArray
        }

    let paragraphsText : Extractor<string> = 
        cellsParagraphsText |>> Common.fromLines

    /// This function matches the regex pattern to the 'inner text'
    /// of the row.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let innerTextIsMatch (pattern:string) : Extractor<bool> = 
        extractor { 
            let! inner = innerText 
            let! regexOpts = getRegexOptions ()
            return Regex.IsMatch( input = inner
                                , pattern = pattern
                                , options = regexOpts )
        }


    let isMatch (cellPatterns:string []) : Extractor<bool> = 
        extractor { 
            let! arrCells = cells |>> Seq.toArray
            let! pairs = 
                liftAction "zip mismatch" (fun _ -> Array.zip cellPatterns arrCells) |>> Array.toList
            return! forallM (fun (patt,cell1) -> focus cell1 (Cell.isMatch patt)) pairs
        }

    // TODO - use regex groups for a function like rowIsMatch that returns matches

    /// Use with caution - a `RegularExpressions.Match` has not necessarily 
    /// matched the input string. The `.Success` property may be false.
    let regexMatch (cellPatterns:string []) : Extractor<RegularExpressions.Match []> = 
        extractor { 
            let! arrCells = cells |>> Seq.toArray
            let! pairs = 
                liftAction "zip mismatch" (fun _ -> Array.zip cellPatterns arrCells) |>> Array.toList
            return! mapM (fun (patt,cell1) -> focus cell1 (Cell.regexMatch patt)) pairs |>> List.toArray
        }


    /// Prefer this to `regexMatch` if you are expecting an array 
    /// of successful matches and you don't need to inspect the result
    /// (e.g. for a match group).
    let regexMatchValues (cellPatterns:string []) : Extractor<string []> = 
        extractor { 
            let! arrCells = cells |>> Seq.toArray
            let! pairs = 
                liftAction "zip mismatch" (fun _ -> Array.zip cellPatterns arrCells) |>> Array.toList
            return! mapM (fun (patt,cell1) -> 
                            focus cell1 (Cell.regexMatchValue patt)) pairs |>> List.toArray
        }

    /// Parse a two column row with "name" in the first cell and 
    /// "value" in the second cell.
    let nameValue2Row (namePattern:string) : Extractor<string> = 
        extractor { 
            let! arr = regexMatchValues [| namePattern; ".*" |]
            let! ans = liftAction "nameValue1Row - bad index" (fun _ -> arr.[1])
            return ans
        }

    /// Parse a three column row with "name" in the first cell and 
    /// "value1" and "value2" in the second and third cells.
    let nameValue3Row (namePattern:string) : Extractor<string * string> = 
        extractor { 
            let! arr = regexMatchValues [| namePattern; ".*"; ".*" |]
            let! ans1 = liftAction "nameValue2Row - bad index" (fun _ -> arr.[1])
            let! ans2 = liftAction "nameValue2Row - bad index" (fun _ -> arr.[2])
            return (ans1, ans2)
        }

    /// Parse a four column row with "name" in the first cell and 
    /// "value1", "value2" and "value3" in the second, third and fourth cells.
    let nameValue4Row (namePattern:string) : Extractor<string * string * string> = 
        extractor { 
            let! arr = regexMatchValues [| namePattern; ".*"; ".*" ; ".*" |]
            let! ans1 = liftAction "nameValue3Row - bad index" (fun _ -> arr.[1])
            let! ans2 = liftAction "nameValue3Row - bad index" (fun _ -> arr.[2])
            let! ans3 = liftAction "nameValue3Row - bad index" (fun _ -> arr.[3])
            return (ans1, ans2, ans3)
        }

    /// TODO - nameValuesRow (namePattern:string) : Extractor<string []> 