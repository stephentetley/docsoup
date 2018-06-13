// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocSoup.Base

open System.IO
open System.Text.RegularExpressions

// Add references via the COM tab for Office and Word
// All the PIA stuff online is outdated for Office 365 / .Net 4.5 / VS2015 
open Microsoft.Office.Interop



let rbox (v : 'a) : obj ref = ref (box v)



// StringReader appears to be the best way of doing this. 
// Trying to split on a character (or character combo e.g. "\r\n") seems unreliable.
let sRestOfLine (s:string) : string = 
    use reader = new StringReader(s)
    reader.ReadLine ()


// Range is a very heavy object to be manipulating start and end points
// Use an alternative...
[<StructuredFormatDisplay("Region: {RegionStart} to {RegionEnd}")>]
type Region = { RegionStart : int; RegionEnd : int}


let extractRegion (range:Word.Range) : Region = { RegionStart = range.Start; RegionEnd = range.End }
    
let maxRegion (doc:Word.Document) : Region = extractRegion <| doc.Range()

let getRange (region:Region) (doc:Word.Document) : Word.Range = 
    doc.Range(rbox <| region.RegionStart, rbox <| region.RegionEnd - 1)

// Use (Single Case) Struct Unions to get the same things as Haskell Newtypes.

let isSubregionOf (haystack:Region) (needle:Region) : bool = 
    needle.RegionStart >= haystack.RegionStart && needle.RegionEnd <= haystack.RegionEnd

let regionText (focus:Region) (doc:Word.Document) : string = 
    let range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
    Regex.Replace(range.Text, @"[\p{C}-[\r\n]]+", "")

let regionPlus (r1:Region) (r2:Region) : Region = 
    { RegionStart = min r1.RegionStart r2.RegionStart
      RegionEnd = max r1.RegionEnd r2.RegionEnd }

let regionConcat (regions:Region list) : Region option = 
    match regions with
    | [] -> None
    | (r::rs) -> Some <| List.fold regionPlus r rs


// cells
let internal softCell (table:Word.Table) (row:int) (col:int) : option<Word.Cell> = 
    try 
        Some <| table.Cell(row, col)
    with
    | _ -> None

let tryFindCell (predicate:Word.Cell -> bool) (table:Word.Table) : option<Word.Cell> = 
    let rowMax = table.Rows.Count 
    let colMax = table.Columns.Count

    let rec work (row:int) (col:int) : option<Word.Cell> = 
        if row >= rowMax then 
            None
        else
            if col >= colMax then 
                work (row+1) 0 
            else
                match softCell table row col with
                | None -> work row (col+1)
                | Some cell ->
                    if predicate cell then
                        Some cell
                    else work row (col+1)
    work 0 0                 



[<Struct>]
type TableAnchor = internal TableAnchor of int
let internal getTableIndex (anchor:TableAnchor) : int = 
    match anchor with | TableAnchor ix -> ix

type CellAnchor = 
    internal 
        { TableIndex: TableAnchor
          RowIndex: int
          ColumnIndex: int }