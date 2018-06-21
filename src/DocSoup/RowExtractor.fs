// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


/// Extract from a single table.
module DocSoup.RowExtractor

open System.Text.RegularExpressions
open Microsoft.Office.Interop
open FParsec

open DocSoup.Base
open DocSoup




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


type RowPhantom = class end
type CellPhantom = class end


/// RowExtractor is intended to be minimal and only run from TablesExtractor
/// RowExtractor is Reader(immutable)+State+Error
type RowExtractor<'nav,'a> = 
    RowExtractor of (Word.Table -> CellIndex -> RowResult<'a>)

type RowParser<'a> = RowExtractor<RowPhantom, 'a>
type CellParser<'a> = RowExtractor<CellPhantom, 'a>


/// This is ignorant of coalesced rows.
let private rowInBounds (table:Word.Table) (ix:CellIndex) : bool = 
    ix.RowIx >= 1 && ix.RowIx <= table.Rows.Count

let private cellInBounds (table:Word.Table) (ix:CellIndex) : bool = 
    ix.ColIx >= 1 && ix.ColIx <= table.Columns.Count

let private cellExists (table:Word.Table) (ix:CellIndex) : bool =
    try 
        let _ = table.Cell(Row = ix.RowIx, Column = ix.ColIx)
        true
    with
    | _ -> false
    

let inline private apply1 (ma: RowExtractor<'nav, 'a>) 
                            (table: Word.Table)
                            (ix: CellIndex) : RowResult<'a>= 
    let (RowExtractor f) = ma in f table ix

let inline rereturn (x:'a) : RowExtractor<'nav,'a> = 
    RowExtractor <| fun _ ix -> ROk(ix,x)


let inline private bindM (ma: RowExtractor<'nav,'a>) 
                            (f: 'a -> RowExtractor<'nav,'b>) : RowExtractor<'nav,'b> =
    RowExtractor <| fun table ix -> 
        match apply1 ma table ix with
        | RErr msg -> RErr msg
        | ROk (ix1,a) -> apply1 (f a) table ix1


let inline rezero () : RowExtractor<'nav,'a> = 
    RowExtractor <| fun _ _ -> RErr "rzero"


let inline private combineM (ma:RowExtractor<'nav,unit>) 
                                (mb:RowExtractor<'nav,unit>) : RowExtractor<'nav,unit> = 
    RowExtractor <| fun table ix -> 
        match apply1 ma table ix with
        | RErr msg -> RErr msg
        | ROk (ix1,a) -> apply1 mb table ix1


let inline private  delayM (fn:unit -> RowExtractor<'nav,'a>) : RowExtractor<'nav,'a> = 
    bindM (rereturn ()) fn 




type RowExtractorBuilder () = 
    member self.Return x            = rereturn x
    member self.Bind (p,f)          = bindM p f
    member self.Zero ()             = rezero ()
    member self.Combine (ma,mb)     = combineM ma mb
    member self.Delay fn            = delayM fn


// Prefer "parse" to "parser" for the _Builder instance

let (parseRows:RowExtractorBuilder) = new RowExtractorBuilder ()



let (&>>>=) (ma:RowExtractor<'nav,'a>) 
            (fn:'a -> RowExtractor<'nav,'b>) : RowExtractor<'nav,'b> = 
    bindM ma fn



// *************************************
// Errors

let rowError (msg:string) : RowExtractor<'nav,'a> = 
    RowExtractor <| fun _ _ -> RErr msg

let swapRowError (msg:string) (ma:RowExtractor<'nav,'a>) : RowExtractor<'nav,'a> = 
    RowExtractor <| fun table ix ->
        match apply1 ma table ix with
        | RErr _ -> RErr msg
        | ROk (ix1,a) -> ROk (ix1,a)


let (<&??>) (ma:RowExtractor<'nav,'a>) (msg:string) : RowExtractor<'nav,'a> = 
    swapRowError msg ma

let (<??&>) (msg:string) (ma:RowExtractor<'nav,'a>) : RowExtractor<'nav,'a> = 
    swapRowError msg ma



// API issue
// Can all fmapM like things be done the outer monad (i.e. DocExtractor)?
// ...
// Not really, it feels like they are essential. So we will have to rely 
// on qualified names and "respellings" for the operators.



let fmapM (fn:'a -> 'b) (ma:RowExtractor<'nav,'a>) : RowExtractor<'nav,'b> = 
    RowExtractor <| fun table ix -> 
       match apply1 ma table ix with
       | RErr msg -> RErr msg
       | ROk (ix1,a) -> ROk (ix1, fn a)


/// Operator for fmap.
let (&|>>>) (ma:RowExtractor<'nav,'a>) (fn:'a -> 'b) : RowExtractor<'nav,'b> = 
    fmapM fn ma

/// Flipped fmap.
let (<<<|&) (fn:'a -> 'b) (ma:RowExtractor<'nav,'a>) : RowExtractor<'nav,'b> = 
    fmapM fn ma


let mapM (p: 'a -> RowExtractor<'nav,'b>) (source:'a list) : RowExtractor<'nav,'b list> = 
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
let (<|||>) (ma:RowExtractor<'nav,'a>) 
            (mb:RowExtractor<'nav,'a>) : RowExtractor<'nav,'a> = 
    RowExtractor <| fun table ix -> 
        match apply1 ma table ix with
        | RErr msg -> apply1 mb table ix
        | ROk (ix1,a) -> ROk (ix1,a)


/// Optionally parses. When the parser fails return None and don't move the cursor position.
let optional (ma:RowExtractor<'nav,'a>) : RowExtractor<'nav,'a option> = 
    RowExtractor <| fun table ix ->
        match apply1 ma table ix with
        | RErr _ -> ROk (ix, None)
        | ROk (ix1,a) -> ROk (ix1, Some a)

/// Perform two actions in sequence. Ignore the results of the second action if both succeed.
let seqL (ma:RowExtractor<'nav,'a>) 
            (mb:RowExtractor<'nav,'b>) : RowExtractor<'nav,'a> = 
    parseRows { 
        let! a = ma
        let! _ = mb
        return a
    }

/// Perform two actions in sequence. Ignore the results of the first action if both succeed.
let seqR (ma:RowExtractor<'nav,'a>) 
            (mb:RowExtractor<'nav,'b>) : RowExtractor<'nav,'b> = 
    parseRows { 
        let! _ = ma
        let! b = mb
        return b
    }

let (.&>>>) (ma:RowExtractor<'nav,'a>) 
                (mb:RowExtractor<'nav,'b>) : RowExtractor<'nav,'a> = 
    seqL ma mb

let (&>>>.) (ma:RowExtractor<'nav,'a>) 
                (mb:RowExtractor<'nav,'b>) : RowExtractor<'nav,'b> = 
    seqR ma mb

/// Test the parser at the current position, if the parser succeeds 
/// return its answer but don't move forward.
/// If the parser fails - lookahead fails.
let lookahead (ma:RowExtractor<'nav,'a>) : RowExtractor<'nav,'a> = 
    RowExtractor <| fun table ix ->
        match apply1 ma table ix with
        | RErr msg -> RErr msg
        | ROk (_,a) -> ROk (ix,a)


    

// *************************************
// Run function

/// Run a RowExtractor. 
let runRowExtractor (ma:RowExtractor<'nav,'a>) (table:Word.Table) : RowResult<'a> =
    try 
        let ix = CellIndex.First
        apply1 ma table ix
    with
    | FatalParseError msg -> raise (FatalParseError msg)
    | _ -> RErr "runRowExtractor"



// *************************************
// Get cell text 


let private cellRange : CellParser<Word.Range> = 
    RowExtractor <| fun table ix -> 
        try 
            let cell : Word.Cell = 
                table.Cell(Row = ix.RowIx, Column = ix.ColIx)
            let ix1 = { ix with ColIx = ix.ColIx+1 }
            ROk (ix1, cell.Range)
        with
        | FatalParseError msg -> raise (FatalParseError msg)
        | _ -> RErr "cellRange"



let cellText : CellParser<string> = 
    swapRowError "cellText" <| fmapM cleanRangeText cellRange



// *************************************
// Navigation

let row (ma:CellParser<'a>) : RowParser<'a> = 
    RowExtractor <| fun table ix -> 
        try
            match apply1 ma table ix with
            | RErr msg -> RErr msg
            | ROk (_,a) -> 
                let ix1 = { RowIx = ix.RowIx+1; ColIx = 1 }
                ROk (ix1, a)
        with
        | FatalParseError msg -> raise (FatalParseError msg)
        | _ -> RErr "row"

let skipCell : CellParser<unit> = 
    RowExtractor <| fun _ ix -> 
        let ix1 = ix.IncrCol
        ROk (ix1, ())

let skipRow : RowParser<unit> = 
    RowExtractor <| fun _ ix -> 
        let ix1 = { RowIx = ix.RowIx + 1; ColIx = 1 }
        ROk (ix1, ())


/// This should continue to end of row.
/// It should also be imprevious to coalesced cells
let manyCells (ma:CellParser<'a>) : CellParser<'a list> = 
    RowExtractor <| fun table ix0 -> 
        try
            let rec work ix ac = 
                if cellInBounds table ix then 
                    // Consider whether cell is coalesced... 
                    if cellExists table ix then 
                        match apply1 ma table ix with
                        | RErr msg -> ROk (ix, List.rev ac)
                        | ROk (ix1,a) -> work ix1 (a::ac)
                    else
                        work ix.IncrCol ac
                else
                    ROk (ix, List.rev ac) 
            work ix0 []
        with
        | FatalParseError msg -> raise (FatalParseError msg)
        | _ -> RErr "manyCells"

/// This should continue to end of table.
let manyRows (ma:RowParser<'a>) : RowParser<'a list> = 
    RowExtractor <| fun table ix0 -> 
        try
            let rec work ix ac = 
                if rowInBounds table ix then 
                    match apply1 ma table ix with
                    | RErr msg -> ROk (ix, List.rev ac)
                    | ROk (ix1,a) -> 
                        work ix1 (a::ac)
                else
                    ROk (ix, List.rev ac)
            work ix0 []
        with
        | FatalParseError msg -> raise (FatalParseError msg)
        | _ -> RErr "manyRows"

/// This should tolerate coalesced cells.
let manyRowsTill (ma:RowParser<'a>) 
                    (endP:RowParser<_>) : RowParser<'a list> = 
    RowExtractor <| fun table ix0 -> 
        try
            let rec work ix ac = 
                if rowInBounds table ix then 
                    match apply1 endP table ix with
                    | RErr _ -> 
                        match apply1 ma table ix with
                        | RErr msg -> RErr msg
                        | ROk (ix1,a) -> work ix1 (a::ac)
                    | ROk (ix1,_) -> ROk (ix1, List.rev ac)
                else
                    RErr "rowInBounds (end of table)"
            work ix0 []
        with
        | FatalParseError msg -> raise (FatalParseError msg)
        | _ -> RErr "rowInBounds"


/// This should tolerate coalesced cells.
let manyCellsTill (ma:CellParser<'a>) 
                    (endP:CellParser<_>) : CellParser<'a list> = 
    RowExtractor <| fun table ix0 -> 
        try
            let rec work ix ac = 
                if cellInBounds table ix then 
                    // Consider whether cell is coalesced... 
                    if cellExists table ix then 
                        match apply1 endP table ix with
                        | RErr _ -> 
                            match apply1 ma table ix with
                            | RErr msg -> RErr msg
                            | ROk (ix1,a) -> work ix1 (a::ac)
                        | ROk (ix1,_) -> ROk (ix1, List.rev ac)
                    else
                        work ix.IncrCol ac
                else
                    RErr "manyCellsTill (end of row)"
            work ix0 []
        with
        | FatalParseError msg -> raise (FatalParseError msg)
        | _ -> RErr "manyCellsTill"

let skipRowsTill (ma:RowParser<'a>) : RowParser<'a> = 
    manyRowsTill skipRow (lookahead ma) &>>>. ma


let skipCellsTill (ma:CellParser<'a>) : CellParser<'a> = 
    manyCellsTill skipCell (lookahead ma) &>>>. ma


// *************************************
// Debugging 

let printIx () : RowExtractor<'nav,unit> = 
    RowExtractor <| fun _ ix ->
        printfn "Index: Row=%i, Col=%i" ix.RowIx ix.ColIx
        ROk (ix, ())



let fatal (msg1:string) (ma:RowExtractor<'nav,'a>) : RowExtractor<'nav,'a> =
    RowExtractor <| fun table ix ->
        match apply1 ma table ix with
        | ROk (ix1,a) -> ROk (ix1,a)
        | RErr msg -> 
            let text = sprintf "FATAL (%s):\n%s" msg1 msg
            raise (FatalParseError text)


// *************************************
// Metric info

let getTableDimensions : RowParser<int * int> = 
    RowExtractor <| fun table ix ->
        ROk (ix, (table.Rows.Count, table.Columns.Count))
        
/// Cell count (columns) of current row.
let getCellCount : CellParser<int> = 
    RowExtractor <| fun table ix -> 
        try 
            let rowCells:Word.Cells = table.Rows.Item(ix.RowIx).Cells
            ROk (ix, rowCells.Count)
        with
        | FatalParseError msg -> raise (FatalParseError msg)
        | _ -> RErr "getCellCount"


// *************************************
// Assert        


let assertInBounds () : RowExtractor<'nav,unit> = 
    RowExtractor <| fun table ix ->
        if (ix.RowIx >= 1 && ix.RowIx <= table.Rows.Count) &&
            (ix.ColIx >= 1 && ix.ColIx <= table.Columns.Count) then 
            ROk (ix, ())
        else
            RErr "assertInBounds"


let internal assertCellTest 
                (test:string -> bool) 
                (failCont:string ->CellParser<_>) : CellParser<unit> = 
    cellText &>>>= fun str ->
    if test str then 
        rereturn ()
    else
        failCont str

let assertCellText (str:string) : CellParser<unit> = 
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellText failed - found '%s'; expecting '%s'" cellText str
        rowError msg
    assertCellTest (fun s -> str.Equals(s)) errCont


/// Assert the current cell matches the pattern
/// This is a full Regexp pattern not a Word pattern
let assertCellRegex (regexpPattern:string) : CellParser<unit> = 
    let matchProc (str:string) = Regex.Match(str, regexpPattern).Success
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellMatches failed - found '%s'; expecting match on '%s'" cellText regexpPattern
        rowError msg
    assertCellTest matchProc errCont

let assertCellWordMatch (wordPattern:string) : CellParser<unit> = 
    cellRange &>>>= fun (rng:Word.Range) -> 
        match boundedFindPattern1 wordPattern id rng with
        | Some _ -> rereturn ()
        | None -> rowError "assertCellWordMatch"


let assertCellEmpty : CellParser<unit> = 
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellEmpty failed - found '%s'" cellText
        rowError msg
    assertCellTest (fun str -> str.Length = 0) errCont


let assertCellTextNot (str:string) : CellParser<unit> = 
    let errCont (cellText:string) = 
        let msg = sprintf "assertCellText failed - found '%s'; expecting '%s'" cellText str
        rowError msg
    assertCellTest (fun s -> not <| str.Equals(s)) errCont



// *************************************
// String level parsing with FParsec

/// We expect string level parsers might fail. 
/// Use this with caution or use execFParsecFallback.
let cellParse (parser:ParsecParser<'a>) : CellParser<'a> = 
    cellText &>>>= fun text -> 
        let name = "none" 
        match runParserOnString parser () name text with
        | Success(ans,_,_) -> rereturn ans
        | Failure(msg,_,_) -> rowError msg
    


// Returns fallback text if FParsec fails.
let cellParseFallback (parser:ParsecParser<'a>) : CellParser<FParsecFallback<'a>> = 
    cellText &>>>= fun text -> 
        let name = "none" 
        match runParserOnString parser () name text with
        | Success(ans,_,_) -> rereturn (FParsecOk ans)
        | Failure(msg,_,_) -> rereturn (FallbackText text)



// *************************************
// Introspect the position

/// row * column
let askCellPosition : CellParser<int * int> = 
    RowExtractor <| fun _ ix -> ROk (ix, (ix.RowIx, ix.ColIx))



/// Note row "width" is dynamic, this means the count acknowledges 
/// coalesced cells
let endOfRow : CellParser<unit> = 
    RowExtractor <| fun table ix -> 
        try 
            let rowCells:Word.Cells = table.Rows.Item(ix.RowIx).Cells
            let cellCount = rowCells.Count
            if ix.ColIx > cellCount then 
                ROk (ix, ())
            else
                RErr "endOfRow (not at end)"
        with
        | FatalParseError msg -> raise (FatalParseError msg)
        | _ -> RErr "endOfRow (system failure)"

let endOfTable : RowParser<unit> = 
    RowExtractor <| fun table ix -> 
        try
            let rowCells:Word.Cells = table.Rows.Item(ix.RowIx).Cells
            let cellCount = rowCells.Count
            if ix.RowIx > table.Rows.Count + 1 then 
                ROk (ix, ())
            else
                RErr "endOfTable (not at end)"
        with
        | FatalParseError msg -> raise (FatalParseError msg)
        | _ -> RErr "endOfTable (system failure)"
