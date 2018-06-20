// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


/// Extract from a single table.
module DocSoup.RowExtractor

open System.Text.RegularExpressions
open Microsoft.Office.Interop
open FParsec

open DocSoup.Base

(* 
/// This is probably the best we can get without tricks.
/// Ideally we wouldn't need ``row``, but I can't see how to avoid it.
/// With a phantom type, we can likely put useful constraints on 
/// composing parsers.
let parser : RowExtractor<SizeRecord> = 
    rowParser {
        do! row (assertString "Table Name")                         /// cell  (1,1)
        let! height = row (assertMatch "*Height" >>>. cellText)     /// cells (2,1) & (2,2)
        let! width  = row (assertMatch "*Width"  >>>. cellText)     /// cells (3,1) & (3,2)
        return { 
            Width = width
            Height = height }
    }
*)

type RowResult<'a> = 
    | RErr of string
    | ROk of CellIndex * 'a




/// RowExtractor is intended to be minimal and only run from TablesExtractor
/// RowExtractor is Reader(immutable)+State+Error
type RowExtractor<'a> = 
    RowExtractor of (Word.Table -> CellIndex -> RowResult<'a>)



let inline private apply1 (ma: RowExtractor<'a>) 
                            (table: Word.Table)
                            (ix: CellIndex) : RowResult<'a>= 
    let (RowExtractor f) = ma in f table ix

let inline rereturn (x:'a) : RowExtractor<'a> = 
    RowExtractor <| fun _ ix -> ROk(ix,x)


let inline private bindM (ma: RowExtractor<'a>) 
                            (f: 'a -> RowExtractor<'b>) : RowExtractor<'b> =
    RowExtractor <| fun table ix -> 
        match apply1 ma table ix with
        | RErr msg -> RErr msg
        | ROk (ix1,a) -> apply1 (f a) table ix1


let inline rezero () : RowExtractor<'a> = 
    RowExtractor <| fun _ _ -> RErr "rzero"


let inline private combineM (ma:RowExtractor<unit>) 
                                (mb:RowExtractor<unit>) : RowExtractor<unit> = 
    RowExtractor <| fun table ix -> 
        match apply1 ma table ix with
        | RErr msg -> RErr msg
        | ROk (ix1,a) -> apply1 mb table ix1


let inline private  delayM (fn:unit -> RowExtractor<'a>) : RowExtractor<'a> = 
    bindM (rereturn ()) fn 




type RowExtractorBuilder() = 
    member self.Return x            = rereturn x
    member self.Bind (p,f)          = bindM p f
    member self.Zero ()             = rezero ()
    member self.Combine (ma,mb)     = combineM ma mb
    member self.Delay fn            = delayM fn

// Prefer "parse" to "parser" for the _Builder instance

let (parseRows:RowExtractorBuilder) = new RowExtractorBuilder()



let (&>>=) (ma:RowExtractor<'a>) 
            (fn:'a -> RowExtractor<'b>) : RowExtractor<'b> = 
    bindM ma fn



// *************************************
// Errors

let rowError (msg:string) : RowExtractor<'a> = 
    RowExtractor <| fun _ _ -> RErr msg

let swapRowError (msg:string) (ma:RowExtractor<'a>) : RowExtractor<'a> = 
    RowExtractor <| fun table ix ->
        match apply1 ma table ix with
        | RErr _ -> RErr msg
        | ROk (ix1,a) -> ROk (ix1,a)


let (<&??>) (ma:RowExtractor<'a>) (msg:string) : RowExtractor<'a> = 
    swapRowError msg ma

let (<??&>) (msg:string) (ma:RowExtractor<'a>) : RowExtractor<'a> = 
    swapRowError msg ma



// API issue
// Can all fmapM like things be done the outer monad (i.e. DocExtractor)?
// ...
// Not really, it feels like they are essential. So we will have to rely 
// on qualified names and "respellings" for the operators.



let fmapM (fn:'a -> 'b) (ma:RowExtractor<'a>) : RowExtractor<'b> = 
    RowExtractor <| fun table ix -> 
       match apply1 ma table ix with
       | RErr msg -> RErr msg
       | ROk (ix1,a) -> ROk (ix1, fn a)


/// Operator for fmap.
let (&|>>>) (ma:RowExtractor<'a>) (fn:'a -> 'b) : RowExtractor<'b> = 
    fmapM fn ma

/// Flipped fmap.
let (<<<|&) (fn:'a -> 'b) (ma:RowExtractor<'a>) : RowExtractor<'b> = 
    fmapM fn ma


let mapM (p: 'a -> RowExtractor<'b>) (source:'a list) : RowExtractor<'b list> = 
    RowExtractor <| fun table ix0 -> 
        let rec work ix ac ys = 
            match ys with
            | [] -> ROk (ix, List.rev ac)
            | z :: zs -> 
                match apply1 (p z) table ix with
                | RErr msg -> RErr msg
                | ROk (ix1,ans) -> work ix1 (ans::ac) zs
        work ix0 [] source




/// Left biased choice
let (<|||>) (ma:RowExtractor<'a>) (mb:RowExtractor<'a>) : RowExtractor<'a> = 
    RowExtractor <| fun table ix -> 
        match apply1 ma table ix with
        | RErr msg -> apply1 mb table ix
        | ROk (ix1,a) -> ROk (ix1,a)


/// Optionally parses. When the parser fails return None and don't move the cursor position.
let optional (ma:RowExtractor<'a>) : RowExtractor<'a option> = 
    RowExtractor <| fun table ix ->
        match apply1 ma table ix with
        | RErr _ -> ROk (ix, None)
        | ROk (ix1,a) -> ROk (ix1, Some a)

/// Perform two actions in sequence. Ignore the results of the second action if both succeed.
let seqL (ma:RowExtractor<'a>) (mb:RowExtractor<'b>) : RowExtractor<'a> = 
    parseRows { 
        let! a = ma
        let! _ = mb
        return a
    }

/// Perform two actions in sequence. Ignore the results of the first action if both succeed.
let seqR (ma:RowExtractor<'a>) (mb:RowExtractor<'b>) : RowExtractor<'b> = 
    parseRows { 
        let! _ = ma
        let! b = mb
        return b
    }

let (.&>>>) (ma:RowExtractor<'a>) (mb:RowExtractor<'b>) : RowExtractor<'a> = 
    seqL ma mb

let (&>>>.) (ma:RowExtractor<'a>) (mb:RowExtractor<'b>) : RowExtractor<'b> = 
    seqR ma mb


// *************************************
// Run function

/// Run a RowExtractor. 
let runRowExtractor (ma:RowExtractor<'a>) (table:Word.Table) : RowResult<'a> =
    try 
        let ix = CellIndex.First
        apply1 ma table ix
    with
    | _ -> RErr "runRowExtractor"


// *************************************
// Navigation

let row (ma:RowExtractor<'a>) : RowExtractor<'a> = 
    RowExtractor <| fun table ix -> 
        try
            match apply1 ma table ix with
            | RErr msg -> RErr msg
            | ROk (_,a) -> 
                let ix1 = { RowIx = ix.RowIx+1; ColIx = 1 }
                ROk (ix1, a)
        with
        | _ -> RErr "row"

let cellAnything : RowExtractor<string> = 
    RowExtractor <| fun table ix -> 
        try 
            let cell : Word.Cell = 
                table.Cell(Row = ix.RowIx, Column = ix.ColIx)
            let text = cleanRangeText cell.Range
            let ix1 = { ix with ColIx = ix.ColIx+1 }
            ROk (ix1, text)
        with
        | _ -> RErr "cellAnything"



// *************************************
// Metric info

let getTableDimensions : RowExtractor<int * int> = 
    RowExtractor <| fun table ix ->
        ROk (ix, (table.Rows.Count, table.Columns.Count))
        
/// Cell count (coulmns) of current row.
let getCellCount : RowExtractor<int> = 
    RowExtractor <| fun table ix -> 
        try 
            let rowCells:Word.Cells = table.Rows.Item(ix.RowIx).Cells
            ROk (ix, rowCells.Count)
        with
        | _ -> RErr "getCellCount"


// *************************************
// Assert        


let assertInBounds  : RowExtractor<unit> = 
    RowExtractor <| fun table ix ->
        if (ix.RowIx >= 1 && ix.RowIx <= table.Rows.Count) &&
            (ix.ColIx >= 1 && ix.ColIx <= table.Columns.Count) then 
            ROk (ix, ())
        else
            RErr "assertInBounds"


let internal assertCellTest 
                (test:string -> bool) 
                (failCont:string ->RowExtractor<_>) : RowExtractor<unit> = 
    cellAnything &>>= fun str ->
    if test str then 
        rereturn ()
    else
        failCont str

let assertCellText (str:string) : RowExtractor<unit> = 
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellText failed - found '%s'; expecting '%s'" cellText str
        rowError msg
    assertCellTest (fun s -> str.Equals(s)) errCont

let assertCellMatches (pattern:string) : RowExtractor<unit> = 
    let matchProc (str:string) = Regex.Match(str, pattern).Success
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellMatches failed - found '%s'; expecting match on '%s'" cellText pattern
        rowError msg
    assertCellTest matchProc errCont

let assertCellEmpty : RowExtractor<unit> = 
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellEmpty failed - found '%s'" cellText
        rowError msg
    assertCellTest (fun str -> str.Length = 0) errCont

let assertCellTextNot (str:string) : RowExtractor<unit> = 
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellText failed - found '%s'; expecting '%s'" cellText str
        rowError msg
    assertCellTest (fun s -> not <| str.Equals(s)) errCont



// *************************************
// String level parsing with FParsec

/// We expect string level parsers might fail. 
/// Use this with caution or use execFParsecFallback.
let cellParse (parser:ParsecParser<'a>) : RowExtractor<'a> = 
    cellAnything &>>= fun text -> 
        let name = "none" 
        match runParserOnString parser () name text with
        | Success(ans,_,_) -> rereturn ans
        | Failure(msg,_,_) -> rowError msg
    


// Returns fallback text if FParsec fails.
let cellParseFallback (parser:ParsecParser<'a>) : RowExtractor<FParsecFallback<'a>> = 
    cellAnything &>>= fun text -> 
        let name = "none" 
        match runParserOnString parser () name text with
        | Success(ans,_,_) -> rereturn (FParsecOk ans)
        | Failure(msg,_,_) -> rereturn (FallbackText text)



// *************************************
// Introspect the position

/// row * column
let askCellPosition : RowExtractor<int * int> = 
    RowExtractor <| fun _ ix -> ROk (ix, (ix.RowIx, ix.ColIx))



/// Note row "width" is dynamic, this means the count acknowledges 
/// coalesced cells
let endOfRow : RowExtractor<unit> = 
    RowExtractor <| fun table ix -> 
        try 
            let rowCells:Word.Cells = table.Rows.Item(ix.RowIx).Cells
            let cellCount = rowCells.Count
            if ix.ColIx > cellCount + 1 then 
                ROk (ix, ())
            else
                RErr "endOfRow (not at end)"
        with
        | _ -> RErr "endOfRow (system failure)"

let endOfTable : RowExtractor<unit> = 
    RowExtractor <| fun table ix -> 
        try
            let rowCells:Word.Cells = table.Rows.Item(ix.RowIx).Cells
            let cellCount = rowCells.Count
            if ix.RowIx > table.Rows.Count + 1 then 
                ROk (ix, ())
            else
                RErr "endOfTable (not at end)"
        with
        | _ -> RErr "endOfTable (system failure)"
