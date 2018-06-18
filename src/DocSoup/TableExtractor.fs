// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocSoup.TableExtractor

open Microsoft.Office.Interop
open FParsec

open DocSoup.Base


type TResult<'a> = 
    | TErr of string
    | TOk of 'a

let private resultConcat (source:TResult<'a> list) : TResult<'a list> = 
    let rec work ac xs = 
        match xs with
        | [] -> TOk <| List.rev ac
        | TOk a::ys -> work (a::ac) ys
        | TErr msg :: _ -> TErr msg
    work [] source



/// TableExtractor is intended to be minimal and only run 
/// from DocExtractor
/// TableExtractor is Reader(immutable)+Reader+Error
type TableExtractor<'a> = 
    TableExtractor of (Word.Table -> CellIndex -> TResult<'a>)



let inline private apply1 (ma: TableExtractor<'a>) 
                            (table: Word.Table)
                            (pos: CellIndex) : TResult<'a>= 
    let (TableExtractor f) = ma in f table pos

let inline treturn (x:'a) : TableExtractor<'a> = 
    TableExtractor <| fun _ _ -> TOk x


let inline private bindM (ma: TableExtractor<'a>) 
                            (f: 'a -> TableExtractor<'b>) : TableExtractor<'b> =
    TableExtractor <| fun table pos -> 
        match apply1 ma table pos with
        | TErr msg -> TErr msg
        | TOk a -> apply1 (f a) table pos

let inline tzero () : TableExtractor<'a> = 
    TableExtractor <| fun _ _ -> TErr "tzero"


let inline private combineM (ma:TableExtractor<unit>) 
                                (mb:TableExtractor<unit>) : TableExtractor<unit> = 
    TableExtractor <| fun table pos -> 
        match apply1 ma table pos with
        | TErr msg -> TErr msg
        | TOk a -> apply1 mb table pos


let inline private  delayM (fn:unit -> TableExtractor<'a>) : TableExtractor<'a> = 
    bindM (treturn ()) fn 




type TableExtractorBuilder() = 
    member self.Return x            = treturn x
    member self.Bind (p,f)          = bindM p f
    member self.Zero ()             = tzero ()
    member self.Combine (ma,mb)     = combineM ma mb
    member self.Delay fn            = delayM fn

// Prefer "parse" to "parser" for the _Builder instance

let (tableExtract:TableExtractorBuilder) = new TableExtractorBuilder()



let (&>>=) (ma:TableExtractor<'a>) 
            (fn:'a -> TableExtractor<'b>) : TableExtractor<'b> = 
    bindM ma fn


// API issue
// Can all fmapM like things be done the outer monad (i.e. DocExtractor)?
// Not really, it feels like they are essential. So we will have to rely 
// on qualified names and "respellings" for the operators.



let fmapM (fn:'a -> 'b) (ma:TableExtractor<'a>) : TableExtractor<'b> = 
    TableExtractor <| fun table pos -> 
       match apply1 ma table pos with
       | TErr msg -> TErr msg
       | TOk a -> TOk (fn a)


let mapM (p: 'a -> TableExtractor<'b>) (source:'a list) : TableExtractor<'b list> = 
    TableExtractor <| fun table pos -> 
        let rec work ac ys = 
            match ys with
            | [] -> TOk (List.rev ac)
            | z :: zs -> 
                match apply1 (p z) table pos with
                | TErr msg -> TErr msg
                | TOk ans -> work (ans::ac) zs
        work [] source

/// Operator for fmap.
let (||>>>) (ma:TableExtractor<'a>) (fn:'a -> 'b) : TableExtractor<'b> = 
    fmapM fn ma

/// Flipped fmap.
let (<<<||) (fn:'a -> 'b) (ma:TableExtractor<'a>) : TableExtractor<'b> = 
    fmapM fn ma



/// Left biased choice
let (<|||>) (ma:TableExtractor<'a>) (mb:TableExtractor<'a>) : TableExtractor<'a> = 
    TableExtractor <| fun table pos -> 
        match apply1 ma table pos with
        | TErr msg -> apply1 mb table pos
        | TOk a -> TOk a


/// Optionally parses. When the parser fails return None and don't move the cursor position.
let optional (ma:TableExtractor<'a>) : TableExtractor<'a option> = 
    TableExtractor <| fun table pos ->
        match apply1 ma table pos with
        | TErr _ -> TOk None
        | TOk a -> TOk (Some a)


// *************************************
// Run function

// Run a TableExtractor. 
let runTableExtractor (ma:TableExtractor<'a>) (table:Word.Table) : TResult<'a> =
    try 
        let pos = CellIndex.First
        apply1 ma table pos
    with
    | _ -> TErr "runTableExtractor"


// *************************************
// Errors

let tableError (msg:string) : TableExtractor<'a> = 
    TableExtractor <| fun _ _ -> TErr msg

let swapTableError (msg:string) (ma:TableExtractor<'a>) : TableExtractor<'a> = 
    TableExtractor <| fun table pos ->
        match apply1 ma table pos with
        | TErr _ -> TErr msg
        | TOk a -> TOk a

let assertCellInBounds (cell:CellIndex) : TableExtractor<unit> = 
    TableExtractor <| fun table pos ->
        if (cell.RowIx >= 1 && cell.RowIx <= table.Rows.Count) &&
            (cell.ColumnIx >= 1 && cell.ColumnIx <= table.Columns.Count) then 
            TOk ()
        else
            TErr "assertCellInBounds"

// *************************************
// Retrieve input          

let getTableText : TableExtractor<string> = 
    TableExtractor <| fun table _ ->
        TOk <| cleanRangeText table.Range

let getCellText : TableExtractor<string> = 
    TableExtractor <| fun table pos ->
        try 
            let cell = table.Cell(pos.RowIx, pos.ColumnIx)
            TOk <| cleanRangeText cell.Range
        with
        | _ -> TErr "getCellText"



// *************************************
// String level parsing with FParsec

// We expect string level parsers might fail. 
// Use this with caution or use execFParsecFallback.
let cellParse (parser:ParsecParser<'a>) : TableExtractor<'a> = 
    tableExtract { 
        let! text = getCellText
        let name = "none" 
        match runParserOnString parser () name text with
        | Success(ans,_,_) -> return ans
        | Failure(msg,_,_) -> tableError msg |> ignore
    }


// Returns fallback text if FParsec fails.
let cellParseFallback (parser:ParsecParser<'a>) : TableExtractor<FParsecFallback<'a>> = 
    tableExtract { 
        let! text = getCellText
        let name = "none" 
        match runParserOnString parser () name text with
        | Success(ans,_,_) -> return (FParsecOk ans)
        | Failure(msg,_,_) -> return (FallbackText text)
    }


// *************************************
// Control the focus


let askCellPosition : TableExtractor<CellIndex> = 
    TableExtractor <| fun _ pos -> TOk pos


/// Restrict focus to a part of the input doc identified by region.
/// Focus type stays the same
let withCell (cell:CellIndex) (ma:TableExtractor<'a>) : TableExtractor<'a> = 
    TableExtractor <| fun table _ -> 
        apply1 ma table cell

// Version of focus that binds the cell returned from a query.
let withCellM (cellQuery:TableExtractor<CellIndex>) 
                (ma:TableExtractor<'a>) : TableExtractor<'a> = 
    cellQuery &>>= fun cell -> withCell cell ma


// *************************************
// Containing cell

/// Return the cell containing needle.
let containingCell (needle:Region) : TableExtractor<CellIndex> = 
    TableExtractor <| fun table _ -> 
        let testCell (cell:Word.Cell) : bool = 
                isSubregionOf (extractRegion cell.Range) needle
  
        match tryFindCell testCell table with 
        | None -> TErr "containingCell - no match"
        | Some cell -> TOk { RowIx = cell.RowIndex; ColumnIx = cell.ColumnIndex }
        



// *************************************
// Search text for "anchors"

/// Find one cell that contains the search string. 
/// This is not guaranteed to be the first cell if there 
/// are multiple matches.
let findCell (search:string) (matchCase:bool) : TableExtractor<CellIndex> =
    TableExtractor <| fun table pos -> 
        match boundedFind1 search matchCase extractRegion table.Range with
        | Some region -> apply1 (containingCell region) table pos
        | None -> TErr <| sprintf "findText - '%s' not found" search

/// Find one cell that matches the pattern. 
/// This is not guaranteed to be the first cell if there 
/// are multiple matches.
let findCellByPattern (search:string) : TableExtractor<CellIndex> =
    TableExtractor <| fun table pos -> 
        match boundedFindPattern1 search extractRegion table.Range with
        | Some region -> apply1 (containingCell region) table pos
        | None -> TErr <| sprintf "findText - '%s' not found" search
        
/// Find all cells that contains the search string.
let findCells (search:string) (matchCase:bool) : TableExtractor<CellIndex list> =
    TableExtractor <| fun table pos -> 
        boundedFindMany search matchCase extractRegion table.Range
            |> List.map (fun region -> apply1 (containingCell region) table pos)
            |> resultConcat



/// Case sensitivity always appears to be true for Wildcard matches.
let findCellsByPattern (search:string) : TableExtractor<CellIndex list> =
    TableExtractor <| fun table pos ->  
        boundedFindPatternMany search extractRegion table.Range
            |> List.map (fun region -> apply1 (containingCell region) table pos)
            |> resultConcat


// *************************************
// Navigation


// Get the cell by index - must be in focus.
// Note - indexing is from 1.
let getCellByIndex (row:int) (col:int) : TableExtractor<CellIndex> = 
    swapTableError "getCellByIndex" <| 
        tableExtract { 
            let cellIx = { RowIx = row; ColumnIx = col }
            let! _ = assertCellInBounds cellIx
            return cellIx
        }



let cellLeft (cell:CellIndex) : TableExtractor<CellIndex> = 
    swapTableError "cellLeft" <| 
        tableExtract { 
            let c1 = cell.DecrCol
            let! _ = assertCellInBounds c1
            return c1
        }


let cellRight (cell:CellIndex) : TableExtractor<CellIndex> = 
    swapTableError "cellRight" <| 
        tableExtract { 
            let c1 = cell.IncrCol
            let! _ = assertCellInBounds c1
            return c1
        }

let cellBelow (cell:CellIndex) : TableExtractor<CellIndex> = 
    swapTableError "cellBelow" <| 
        tableExtract { 
            let c1 = cell.IncrRow
            let! _ = assertCellInBounds c1
            return c1
        }

let cellAbove (cell:CellIndex) : TableExtractor<CellIndex> = 
    swapTableError "cellAbove" <| 
        tableExtract { 
            let c1 = cell.DecrRow
            let! _ = assertCellInBounds c1
            return c1
        }

