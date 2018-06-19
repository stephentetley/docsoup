// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


/// Extract from a single table.
module DocSoup.TableExtractor1

open System.Text.RegularExpressions
open Microsoft.Office.Interop
open FParsec

open DocSoup.Base
open System


type T1Result<'a> = 
    | T1Err of string
    | T1Ok of 'a

let private resultConcat (source:T1Result<'a> list) : T1Result<'a list> = 
    let rec work ac xs = 
        match xs with
        | [] -> T1Ok <| List.rev ac
        | T1Ok a::ys -> work (a::ac) ys
        | T1Err msg :: _ -> T1Err msg
    work [] source



/// TableExtractor1 is intended to be minimal and only run 
/// from DocExtractor
/// TableExtractor1 is Reader(immutable)+Reader+Error
type TableExtractor1<'a> = 
    TableExtractor1 of (Word.Table -> CellIndex -> T1Result<'a>)



let inline private apply1 (ma: TableExtractor1<'a>) 
                            (table: Word.Table)
                            (pos: CellIndex) : T1Result<'a>= 
    let (TableExtractor1 f) = ma in f table pos

let inline t1return (x:'a) : TableExtractor1<'a> = 
    TableExtractor1 <| fun _ _ -> T1Ok x


let inline private bindM (ma: TableExtractor1<'a>) 
                            (f: 'a -> TableExtractor1<'b>) : TableExtractor1<'b> =
    TableExtractor1 <| fun table pos -> 
        match apply1 ma table pos with
        | T1Err msg -> T1Err msg
        | T1Ok a -> apply1 (f a) table pos

let inline t1zero () : TableExtractor1<'a> = 
    TableExtractor1 <| fun _ _ -> T1Err "t1zero"


let inline private combineM (ma:TableExtractor1<unit>) 
                                (mb:TableExtractor1<unit>) : TableExtractor1<unit> = 
    TableExtractor1 <| fun table pos -> 
        match apply1 ma table pos with
        | T1Err msg -> T1Err msg
        | T1Ok a -> apply1 mb table pos


let inline private  delayM (fn:unit -> TableExtractor1<'a>) : TableExtractor1<'a> = 
    bindM (t1return ()) fn 




type TableExtractor1Builder() = 
    member self.Return x            = t1return x
    member self.Bind (p,f)          = bindM p f
    member self.Zero ()             = t1zero ()
    member self.Combine (ma,mb)     = combineM ma mb
    member self.Delay fn            = delayM fn

// Prefer "parse" to "parser" for the _Builder instance

let (table1:TableExtractor1Builder) = new TableExtractor1Builder()



let (&>>=) (ma:TableExtractor1<'a>) 
            (fn:'a -> TableExtractor1<'b>) : TableExtractor1<'b> = 
    bindM ma fn


// API issue
// Can all fmapM like things be done the outer monad (i.e. DocExtractor)?
// ...
// Not really, it feels like they are essential. So we will have to rely 
// on qualified names and "respellings" for the operators.



let fmapM (fn:'a -> 'b) (ma:TableExtractor1<'a>) : TableExtractor1<'b> = 
    TableExtractor1 <| fun table pos -> 
       match apply1 ma table pos with
       | T1Err msg -> T1Err msg
       | T1Ok a -> T1Ok (fn a)


let mapM (p: 'a -> TableExtractor1<'b>) (source:'a list) : TableExtractor1<'b list> = 
    TableExtractor1 <| fun table pos -> 
        let rec work ac ys = 
            match ys with
            | [] -> T1Ok (List.rev ac)
            | z :: zs -> 
                match apply1 (p z) table pos with
                | T1Err msg -> T1Err msg
                | T1Ok ans -> work (ans::ac) zs
        work [] source

/// Operator for fmap.
let (&|>>>) (ma:TableExtractor1<'a>) (fn:'a -> 'b) : TableExtractor1<'b> = 
    fmapM fn ma

/// Flipped fmap.
let (<<<|&) (fn:'a -> 'b) (ma:TableExtractor1<'a>) : TableExtractor1<'b> = 
    fmapM fn ma



/// Left biased choice
let (<|||>) (ma:TableExtractor1<'a>) (mb:TableExtractor1<'a>) : TableExtractor1<'a> = 
    TableExtractor1 <| fun table pos -> 
        match apply1 ma table pos with
        | T1Err msg -> apply1 mb table pos
        | T1Ok a -> T1Ok a


/// Optionally parses. When the parser fails return None and don't move the cursor position.
let optional (ma:TableExtractor1<'a>) : TableExtractor1<'a option> = 
    TableExtractor1 <| fun table pos ->
        match apply1 ma table pos with
        | T1Err _ -> T1Ok None
        | T1Ok a -> T1Ok (Some a)

/// Perform two actions in sequence. Ignore the results of the second action if both succeed.
let seqL (ma:TableExtractor1<'a>) (mb:TableExtractor1<'b>) : TableExtractor1<'a> = 
    table1 { 
        let! a = ma
        let! b = mb
        return a
    }

/// Perform two actions in sequence. Ignore the results of the first action if both succeed.
let seqR (ma:TableExtractor1<'a>) (mb:TableExtractor1<'b>) : TableExtractor1<'b> = 
    table1 { 
        let! a = ma
        let! b = mb
        return b
    }

let (.&>>>) (ma:TableExtractor1<'a>) (mb:TableExtractor1<'b>) : TableExtractor1<'a> = 
    seqL ma mb

let (&>>>.) (ma:TableExtractor1<'a>) (mb:TableExtractor1<'b>) : TableExtractor1<'b> = 
    seqR ma mb


// *************************************
// Run function

// Run a TableExtractor1. 
let runTableExtractor1 (ma:TableExtractor1<'a>) (table:Word.Table) : T1Result<'a> =
    try 
        let pos = CellIndex.First
        apply1 ma table pos
    with
    | _ -> T1Err "runTableExtractor1"


// *************************************
// Errors

let tableError (msg:string) : TableExtractor1<'a> = 
    TableExtractor1 <| fun _ _ -> T1Err msg

let swapTableError (msg:string) (ma:TableExtractor1<'a>) : TableExtractor1<'a> = 
    TableExtractor1 <| fun table pos ->
        match apply1 ma table pos with
        | T1Err _ -> T1Err msg
        | T1Ok a -> T1Ok a

let assertCellInBounds (cell:CellIndex) : TableExtractor1<unit> = 
    TableExtractor1 <| fun table pos ->
        if (cell.RowIx >= 1 && cell.RowIx <= table.Rows.Count) &&
            (cell.ColumnIx >= 1 && cell.ColumnIx <= table.Columns.Count) then 
            T1Ok ()
        else
            T1Err "assertCellInBounds"

// *************************************
// Retrieve input          

let getTableText : TableExtractor1<string> = 
    TableExtractor1 <| fun table _ ->
        T1Ok <| cleanRangeText table.Range

let getCellText : TableExtractor1<string> = 
    TableExtractor1 <| fun table pos ->
        try 
            let cell = table.Cell(pos.RowIx, pos.ColumnIx)
            T1Ok <| cleanRangeText cell.Range
        with
        | _ -> T1Err "getCellText"

// *************************************
// Metric info

let getTableRegion : TableExtractor1<Region> = 
    TableExtractor1 <| fun table pos ->
        try 
            T1Ok <| extractRegion table.Range
        with
        | _ -> T1Err "getTableRegion"


let getCellRegion : TableExtractor1<Region> = 
    TableExtractor1 <| fun table pos ->
        try 
            let cell = table.Cell(pos.RowIx, pos.ColumnIx)
            T1Ok <| extractRegion cell.Range
        with
        | _ -> T1Err "getCellRegion"

// *************************************
// Assert

let internal assertCellTest 
                (test:string -> bool) 
                (failCont:string ->TableExtractor1<_>) : TableExtractor1<unit> = 
    getCellText &>>= fun str ->
    if test str then 
        t1return ()
    else
        failCont str

let assertCellText (str:string) : TableExtractor1<unit> = 
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellText failed - found '%s'; expecting '%s'" cellText str
        tableError msg
    assertCellTest (fun s -> str.Equals(s)) errCont

let assertCellMatches (pattern:string) : TableExtractor1<unit> = 
    let matchProc (str:string) = Regex.Match(str, pattern).Success
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellMatches failed - found '%s'; expecting match on '%s'" cellText pattern
        tableError msg
    assertCellTest matchProc errCont

let assertCellEmpty : TableExtractor1<unit> = 
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellEmpty failed - found '%s'" cellText
        tableError msg
    assertCellTest (fun str -> str.Length = 0) errCont

let assertCellTextNot (str:string) : TableExtractor1<unit> = 
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellText failed - found '%s'; expecting '%s'" cellText str
        tableError msg
    assertCellTest (fun s -> not <| str.Equals(s)) errCont

// *************************************
// String level parsing with FParsec

// We expect string level parsers might fail. 
// Use this with caution or use execFParsecFallback.
let cellParse (parser:ParsecParser<'a>) : TableExtractor1<'a> = 
    table1 { 
        let! text = getCellText
        let name = "none" 
        match runParserOnString parser () name text with
        | Success(ans,_,_) -> return ans
        | Failure(msg,_,_) -> tableError msg |> ignore
    }


// Returns fallback text if FParsec fails.
let cellParseFallback (parser:ParsecParser<'a>) : TableExtractor1<FParsecFallback<'a>> = 
    table1 { 
        let! text = getCellText
        let name = "none" 
        match runParserOnString parser () name text with
        | Success(ans,_,_) -> return (FParsecOk ans)
        | Failure(msg,_,_) -> return (FallbackText text)
    }


// *************************************
// Control the focus


let askCellPosition : TableExtractor1<CellIndex> = 
    TableExtractor1 <| fun _ pos -> T1Ok pos


/// Restrict focus to a part of the input doc identified by region.
/// Focus type stays the same
let withCell (cell:CellIndex) (ma:TableExtractor1<'a>) : TableExtractor1<'a> = 
    TableExtractor1 <| fun table _ -> 
        apply1 ma table cell

// Version of focus that binds the cell returned from a query.
let withCellM (cellQuery:TableExtractor1<CellIndex>) 
                (ma:TableExtractor1<'a>) : TableExtractor1<'a> = 
    cellQuery &>>= fun cell -> withCell cell ma


// *************************************
// Containing cell

/// Return the cell containing needle.
let containingCell (needle:Region) : TableExtractor1<CellIndex> = 
    TableExtractor1 <| fun table _ -> 
        let testCell (cell:Word.Cell) : bool = 
                isSubregionOf (extractRegion cell.Range) needle
  
        match tryFindCell testCell table with 
        | None -> T1Err "containingCell - no match"
        | Some cell -> T1Ok { RowIx = cell.RowIndex; ColumnIx = cell.ColumnIndex }
        



// *************************************
// Search text for "anchors"

/// Find one cell that contains the search string. 
/// This is not guaranteed to be the first cell if there 
/// are multiple matches.
let findCell (search:string) (matchCase:bool) : TableExtractor1<CellIndex> =
    TableExtractor1 <| fun table pos -> 
        match boundedFind1 search matchCase extractRegion table.Range with
        | Some region -> apply1 (containingCell region) table pos
        | None -> T1Err <| sprintf "findText - '%s' not found" search

/// Find one cell that matches the pattern. 
/// This is not guaranteed to be the first cell if there 
/// are multiple matches.
let findCellByPattern (search:string) : TableExtractor1<CellIndex> =
    TableExtractor1 <| fun table pos -> 
        match boundedFindPattern1 search extractRegion table.Range with
        | Some region -> apply1 (containingCell region) table pos
        | None -> T1Err <| sprintf "findText - '%s' not found" search
        
/// Find all cells that contains the search string.
let findCells (search:string) (matchCase:bool) : TableExtractor1<CellIndex list> =
    TableExtractor1 <| fun table pos -> 
        boundedFindMany search matchCase extractRegion table.Range
            |> List.map (fun region -> apply1 (containingCell region) table pos)
            |> resultConcat



/// Case sensitivity always appears to be true for Wildcard matches.
let findCellsByPattern (search:string) : TableExtractor1<CellIndex list> =
    TableExtractor1 <| fun table pos ->  
        boundedFindPatternMany search extractRegion table.Range
            |> List.map (fun region -> apply1 (containingCell region) table pos)
            |> resultConcat


// *************************************
// Navigation


// Get the cell by index - must be in focus.
// Note - indexing is from 1.
let getCellByIndex (row:int) (col:int) : TableExtractor1<CellIndex> = 
    swapTableError "getCellByIndex" <| 
        table1 { 
            let cellIx = { RowIx = row; ColumnIx = col }
            let! _ = assertCellInBounds cellIx
            return cellIx
        }



let cellLeft (cell:CellIndex) : TableExtractor1<CellIndex> = 
    swapTableError "cellLeft" <| 
        table1 { 
            let c1 = cell.DecrCol
            let! _ = assertCellInBounds c1
            return c1
        }


let cellRight (cell:CellIndex) : TableExtractor1<CellIndex> = 
    swapTableError "cellRight" <| 
        table1 { 
            let c1 = cell.IncrCol
            let! _ = assertCellInBounds c1
            return c1
        }

let cellBelow (cell:CellIndex) : TableExtractor1<CellIndex> = 
    swapTableError "cellBelow" <| 
        table1 { 
            let c1 = cell.IncrRow
            let! _ = assertCellInBounds c1
            return c1
        }

let cellAbove (cell:CellIndex) : TableExtractor1<CellIndex> = 
    swapTableError "cellAbove" <| 
        table1 { 
            let c1 = cell.DecrRow
            let! _ = assertCellInBounds c1
            return c1
        }

