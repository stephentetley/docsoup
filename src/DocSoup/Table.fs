// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Table = 
    
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup

    type TableExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Table> 

    let (extractor:TableExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Table>()

    type Extractor<'a> = ExtractMonad<Wordprocessing.Table,'a> 



    let rows : Extractor<seq<Wordprocessing.TableRow>> = 
        asks (fun table -> table.Elements<Wordprocessing.TableRow>())

    let row (index:int) : Extractor<Wordprocessing.TableRow> = 
        extractor { 
            let! xs = rows
            return! liftOption (Seq.tryItem index xs)
        }

    let rowCount : Extractor<int> = rows |>> Seq.length


    let cell (rowIndex:int, columnIndex:int) : Extractor<Wordprocessing.TableCell> = 
        extractor { 
            let! xs = row rowIndex |>> fun r1 -> r1.Elements<Wordprocessing.TableCell>()
            return! liftOption (Seq.tryItem columnIndex xs)
        }

    let firstRow : Extractor<Wordprocessing.TableRow> = row 0  

    let firstCell : Extractor<Wordprocessing.TableCell> = cell (0,0)


    let findRow (predicate:Row.Extractor<bool>) : Extractor<Wordprocessing.TableRow> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            return! findM (fun t1 -> (mreturn t1) &>> predicate) xs
        }

    let findRowIndex (predicate:Row.Extractor<bool>) : Extractor<int> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            return! findIndexM (fun t1 -> (mreturn t1) &>> predicate) xs
        }


    let innerText : Extractor<string> = 
        asks (fun table -> table.InnerText)

    /// This function matches the regex pattern to the 'inner text'
    /// of the table.
    /// The inner text does not preserve whitespace, so **do not**
    /// try to match against a whitespace sensitive pattern.
    let innerTextIsMatch (pattern:string) : Extractor<bool> = 
        extractor { 
            let! inner = innerText 
            return Regex.IsMatch(inner, pattern)
        }


    let findNameValue1Row (namePattern:string) : Extractor<string> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            return! pickM (fun r1 -> (mreturn r1) &>> Row.nameValue1Row namePattern) xs
        }

    let findNameValue2Row (namePattern:string) : Extractor<string * string> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            return! pickM (fun r1 -> (mreturn r1) &>> Row.nameValue2Row namePattern) xs
        }
        
    let findNameValue3Row (namePattern:string) : Extractor<string * string * string> = 
        extractor { 
            let! xs = rows |>> Seq.toList
            return! pickM (fun r1 -> (mreturn r1) &>> Row.nameValue3Row namePattern) xs
        }