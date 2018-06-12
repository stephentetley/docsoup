// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocSoup.DocMonad

open Microsoft.Office.Interop
open FParsec

open DocSoup.Base


type TextParser<'a> = Parser<'a, unit>

type Result<'a> = 
    | Err of string
    | Ok of 'a


let private execFParsec (doc:Word.Document) (region:Region) (p:TextParser<'a>) : Result<'a> = 
    let text = regionText region doc
    let name = doc.Name  
    match runParserOnString p () name text with
    | Success(ans,_,_) -> Ok ans
    | Failure(msg,_,_) -> Err msg


// DocSoup is Reader(immutable)+Reader+Error
type DocSoup<'a> = DocSoup of (Word.Document -> Region -> Result<'a>)


let inline private apply1 (ma : DocSoup<'a>) (doc:Word.Document) (focus:Region) : Result<'a>= 
    let (DocSoup f) = ma in f doc focus

let inline sreturn (x:'a) : DocSoup<'a> = DocSoup <| fun _ _ -> Ok x


let inline private bindM (ma:DocSoup<'a>) (f : 'a -> DocSoup<'b>) : DocSoup<'b> =
    DocSoup <| fun doc focus -> 
        match apply1 ma doc focus with
        | Err msg -> Err msg
        | Ok a -> apply1 (f a) doc focus

let inline szero () : DocSoup<'a> = 
    DocSoup <| fun _ _ -> Err "szero"


let inline private combineM (ma:DocSoup<unit>) (mb:DocSoup<unit>) : DocSoup<unit> = 
    DocSoup <| fun doc focus -> 
        match apply1 ma doc focus with
        | Err msg -> Err msg
        | Ok a -> apply1 mb doc focus

let inline private  delayM (fn:unit -> DocSoup<'a>) : DocSoup<'a> = 
    bindM (sreturn ()) fn 




type DocSoupBuilder() = 
    member self.Return x            = sreturn x
    member self.Bind (p,f)          = bindM p f
    member self.Zero ()             = szero ()
    // member self.For (xs,ma)         = forExprM xs ma
    member self.Combine (ma,mb)     = combineM ma mb
    member self.Delay fn            = delayM fn

 // Prefer "parse" to "parser" for the _Builder instance

let (docSoup:DocSoupBuilder) = new DocSoupBuilder()



let (>>>=) (ma:DocSoup<'a>) (fn:'a -> DocSoup<'b>) : DocSoup<'b> = bindM ma fn


// Common monadic operations
let fmapM (fn:'a -> 'b) (ma:DocSoup<'a>) : DocSoup<'b> = 
    DocSoup <| fun doc focus -> 
       match apply1 ma doc focus with
       | Err msg -> Err msg
       | Ok a-> Ok <| fn a

// This is the nub of embedding FParsec - name clashes.
// We avoid them by using longer names in DocSoup.
let (|>>>) (ma:DocSoup<'a>) (fn:'a -> 'b) : DocSoup<'b> = fmapM fn ma
let (<<<|) (fn:'a -> 'b) (ma:DocSoup<'a>) : DocSoup<'b> = fmapM fn ma

let liftM (fn:'a -> 'x) (ma:DocSoup<'a>) : DocSoup<'x> = fmapM fn ma

let liftM2 (fn:'a -> 'b -> 'x) (ma:DocSoup<'a>) (mb:DocSoup<'b>) : DocSoup<'x> = 
    docSoup { 
        let! a = ma
        let! b = mb
        return (fn a b)
    }

let liftM3 (fn:'a -> 'b -> 'c -> 'x) (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) : DocSoup<'x> = 
    docSoup { 
        let! a = ma
        let! b = mb
        let! c = mc
        return (fn a b c)
    }

let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (md:DocSoup<'d>) : DocSoup<'x> = 
    docSoup { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        return (fn a b c d)
    }


let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (md:DocSoup<'d>) (me:DocSoup<'e>) : DocSoup<'x> = 
    docSoup { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        let! e = me
        return (fn a b c d e)
    }

let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
            (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) 
            (md:DocSoup<'d>) (me:DocSoup<'e>) (mf:DocSoup<'f>) : DocSoup<'x> = 
    docSoup { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        let! e = me
        let! f = mf
        return (fn a b c d e f)
    }

let tupleM2 (ma:DocSoup<'a>) (mb:DocSoup<'b>) : DocSoup<'a * 'b> = 
    liftM2 (fun a b -> (a,b)) ma mb

let tupleM3 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) : DocSoup<'a * 'b * 'c> = 
    liftM3 (fun a b c -> (a,b,c)) ma mb mc

let tupleM4 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (md:DocSoup<'d>) : DocSoup<'a * 'b * 'c * 'd> = 
    liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

let tupleM5 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (md:DocSoup<'d>) (me:DocSoup<'e>) : DocSoup<'a * 'b * 'c * 'd * 'e> = 
    liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

let tupleM6 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (md:DocSoup<'d>) (me:DocSoup<'e>) (mf:DocSoup<'f>) : DocSoup<'a * 'b * 'c * 'd * 'e * 'f> = 
    liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf

let pipeM2 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (fn:'a -> 'b -> 'x) : DocSoup<'x> = 
    liftM2 fn ma mb

let pipeM3 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (fn:'a -> 'b -> 'c -> 'x): DocSoup<'x> = 
    liftM3 fn ma mb mc

let pipeM4 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (md:DocSoup<'d>) (fn:'a -> 'b -> 'c -> 'd -> 'x) : DocSoup<'x> = 
    liftM4 fn ma mb mc md

let pipeM5 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (md:DocSoup<'d>) (me:DocSoup<'e>) (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x): DocSoup<'x> = 
    liftM5 fn ma mb mc md me

let pipeM6 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (md:DocSoup<'d>) (me:DocSoup<'e>) (mf:DocSoup<'f>) (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x): DocSoup<'x> = 
    liftM6 fn ma mb mc md me mf

// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
let alt (ma:DocSoup<'a>) (mb:DocSoup<'a>) : DocSoup<'a> = 
    DocSoup <| fun doc focus ->
        match apply1 ma doc focus with
        | Err _ -> apply1 mb doc focus
        | Ok a -> Ok a

let (<||>) (ma:DocSoup<'a>) (mb:DocSoup<'a>) : DocSoup<'a> = alt ma mb


// Haskell Applicative's (<*>)
let apM (mf:DocSoup<'a ->'b>) (ma:DocSoup<'a>) : DocSoup<'b> = 
    docSoup { 
        let! fn = mf
        let! a = ma
        return (fn a) 
    }

let (<**>) (ma:DocSoup<'a -> 'b>) (mb:DocSoup<'a>) : DocSoup<'b> = apM ma mb

let (<&&>) (fn:'a -> 'b) (ma:DocSoup<'a>) :DocSoup<'b> = fmapM fn ma


// Perform two actions in sequence. Ignore the results of the second action if both succeed.
let seqL (ma:DocSoup<'a>) (mb:DocSoup<'b>) : DocSoup<'a> = 
    docSoup { 
        let! a = ma
        let! b = mb
        return a
    }

// Perform two actions in sequence. Ignore the results of the first action if both succeed.
let seqR (ma:DocSoup<'a>) (mb:DocSoup<'b>) : DocSoup<'b> = 
    docSoup { 
        let! a = ma
        let! b = mb
        return b
    }

let (.>>>) (ma:DocSoup<'a>) (mb:DocSoup<'b>) : DocSoup<'a> = seqL ma mb
let (>>>.) (ma:DocSoup<'a>) (mb:DocSoup<'b>) : DocSoup<'b> = seqR ma mb


let mapM (p: 'a -> DocSoup<'b>) (source:'a list) : DocSoup<'b list> = 
    DocSoup <| fun doc focus -> 
        let rec work ac ys = 
            match ys with
            | [] -> Ok <| List.rev ac
            | z :: zs -> 
                match apply1 (p z) doc focus with
                | Err msg -> Err msg
                | Ok ans -> work (ans::ac) zs
        work [] source

let forM (source:'a list) (p: 'a -> DocSoup<'b>) : DocSoup<'b list> = mapM p source


// Design note - no ``findM`` function
// val findM : test: 'a -> DocSoup<bool> -> source:'a list -> DocSoup<'a>  
// The name "findM" does not tell enough about the function - the crux is
// should Err be an error or re-interpreted as false.



/// A version of findM that finds the first success
/// This allows type changing.
let findSuccessM  (action: 'a -> DocSoup<'b>) (source:'a list) : DocSoup<'b> = 
    DocSoup <| fun doc focus -> 
        let rec work ys = 
            match ys with
            | [] -> Err "findSuccessM - not found"
            | z :: zs -> 
                match apply1 (action z) doc focus with
                | Err _ -> work zs
                | Ok ans -> Ok ans
        work source

let findSuccessesM  (action: 'a -> DocSoup<'b>) (source:'a list) : DocSoup<'b list> = 
    DocSoup <| fun doc focus -> 
        let rec work ac ys = 
            match ys with
            | [] -> Ok <| List.rev ac
            | z :: zs -> 
                match apply1 (action z) doc focus with
                | Err _ -> work ac zs
                | Ok ans -> work (ans::ac) zs
        work [] source

    
let optionToAction (source:option<'a>) (errMsg:string) : DocSoup<'a> = 
    DocSoup <| fun _ _ -> 
        match source with
        | None -> Err errMsg
        | Some a -> Ok a

let optional (ma:DocSoup<'a>) : DocSoup<'a option> = 
    DocSoup <| fun doc focus ->
        match apply1 ma doc focus with
        | Err _ -> Ok None
        | Ok a -> Ok <| Some a


let optionalz (ma:DocSoup<'a>) : DocSoup<unit> = 
    DocSoup <| fun doc focus ->
        match apply1 ma doc focus with
        | Err _ -> Ok ()
        | Ok _ -> Ok ()

/// Turn an operation into a boolean, when the action is success return true 
/// when it fails return false
let boolify (ma:DocSoup<'a>) : DocSoup<bool> = 
    DocSoup <| fun doc focus ->
        match apply1 ma doc focus with
        | Err _ -> Ok false
        | Ok _ -> Ok true


type FallBackResult<'a> = 
    | ParseOk of 'a
    | FallBackText of string

// We expect string level parsers might fail.
// Rather than throw a hard fail, get the source input instead.
let textFallBack (ma:DocSoup<'a>) : DocSoup<FallBackResult<'a>> = 
    DocSoup <| fun doc focus ->
        let text = regionText focus doc
        match apply1 ma doc focus with
        | Err _ -> Ok <| FallBackText text
        | Ok a -> Ok <| ParseOk a

// *************************************
// Run functions

let runOnFile (ma:DocSoup<'a>) (fileName:string) : Result<'a> =
    if System.IO.File.Exists (fileName) then
        let app = new Word.ApplicationClass (Visible = false) :> Word.Application
        try 
            let doc = app.Documents.Open(FileName = ref (fileName :> obj))
            let region1 = maxRegion doc
            let ans = apply1 ma doc region1
            doc.Close(SaveChanges = rbox false)
            app.Quit()
            ans
        with
        | ex -> 
            try 
                app.Quit ()
                Err ex.Message
            with
            | _ -> Err ex.Message
                
    else 
        Err <| sprintf "Cannot find file %s" fileName


let runOnFileE (ma:DocSoup<'a>) (fileName:string) : 'a =
    match runOnFile ma fileName with
    | Err msg -> failwith msg
    | Ok a -> a

let throwError (msg:string) : DocSoup<'a> = 
    DocSoup <| fun _  _ -> Err msg

let swapError (msg:string) (ma:DocSoup<'a>) : DocSoup<'a> = 
    DocSoup <| fun doc focus ->
        match apply1 ma doc focus with
        | Err _ -> Err msg
        | Ok a -> Ok a

let (<&?>) (ma:DocSoup<'a>) (msg:string) : DocSoup<'a> = swapError msg ma


let focus (region:Region) (ma:DocSoup<'a>) : DocSoup<'a> = 
    DocSoup <| fun doc _ -> apply1 ma doc region

let fparse (p:TextParser<'a>) : DocSoup<'a> = 
    DocSoup <| fun doc focus -> execFParsec doc focus p 


/// This gets the text within the current focus!        
let getText : DocSoup<string> =
    DocSoup <| fun doc focus -> 
        let text = regionText focus doc 
        Ok <| text.Trim ()



let findText (search:string) (matchCase:bool) : DocSoup<Region> =
    DocSoup <| fun doc focus  -> 
        let range1 = getRange focus doc
        range1.Find.ClearFormatting ()
        if range1.Find.Execute (FindText = rbox search, 
                                MatchCase = rbox matchCase,
                                MatchWildcards = rbox false,
                                Forward = rbox true) then
            Ok <| extractRegion range1
        else
            Err <| sprintf "findText - '%s' not found" search

/// Case sensitivity always appears to be true for Wildcard matches.
let findPattern (search:string) : DocSoup<Region> =
    DocSoup <| fun doc focus  -> 
        let range1 = getRange focus doc
        range1.Find.ClearFormatting ()
        if range1.Find.Execute (FindText = rbox search, 
                                MatchWildcards = rbox true,
                                MatchCase = rbox true,
                                Forward = rbox true) then
            Ok <| extractRegion range1
        else
            Err <| sprintf "findText - '%s' not found" search

let findTextMany (search:string) (matchCase:bool) : DocSoup<Region list> =
    let rec work (rng:Word.Range) (ac: Region list) : Result<Region list> = 
        rng.Find.Execute (FindText = rbox search, 
                            MatchCase = rbox matchCase,
                            MatchWildcards = rbox false,
                            Forward = rbox true) |> ignore
        if rng.Find.Found then
            let region = extractRegion rng
            work rng (region::ac)
        else
            Ok <| List.rev ac

    DocSoup <| fun doc focus  -> 
        let range1 = getRange focus doc
        range1.Find.ClearFormatting ()
        work range1 []


/// Case sensitivity always appears to be true for Wildcard matches.
let findPatternMany (search:string) : DocSoup<Region list> =
    let rec work (rng:Word.Range) (ac: Region list) : Result<Region list> = 
        rng.Find.Execute (FindText = rbox search, 
                            MatchWildcards = rbox true,
                            MatchCase = rbox true,
                            Forward = rbox true) |> ignore
        if rng.Find.Found then
            let region = extractRegion rng
            work rng (region::ac)
        else
            Ok <| List.rev ac

    DocSoup <| fun doc focus  -> 
        let range1 = getRange focus doc
        range1.Find.ClearFormatting ()
        work range1 []

/// If successful returns the concatenation of all regions.
let findAll (searches:string list) (matchCase:bool) : DocSoup<Region> =
    mapM (fun s -> findText s matchCase) searches >>>= fun xs ->
    optionToAction (regionConcat xs) "findAll - fail" 

/// If successful returns the concatenation of all regions.
let findPatternAll (searches:string list) : DocSoup<Region> =
    mapM findPattern searches >>>= fun xs ->
    optionToAction (regionConcat xs) "findAllPattern - fail" 


let private getTablesInFocus : DocSoup<Word.Table []> = 
    DocSoup <| fun doc focus -> 
        try 
            // O
            let focusTables : Word.Table [] = 
                doc.Range(Start = rbox focus.RegionStart, End = rbox focus.RegionEnd ).Tables |> Seq.cast<Word.Table> |> Seq.toArray
            Ok <| focusTables
        with
        | _ -> Err "getTablesInFocus"


/// Note - the results will be within the current focus!
let tableAreas : DocSoup<Region []> = 
    let extrRegions = Array.map (fun (table:Word.Table) -> table.Range |> extractRegion)
    fmapM extrRegions getTablesInFocus
    



/// Note - the results are indexed within the current focus.
/// Note - this is zero indexed (unlike Word's automation API)
let getTableArea(tid:TableAnchor) : DocSoup<Region> = 
    let index = getTableIndex tid
    tableAreas >>>= fun arr -> 
    try 
        sreturn arr.[index]
    with
    | _ -> throwError 
            (sprintf "getTable - index out of range ix:%i [tables: %i]" index arr.Length)

let tryContainingTable (needle:Region) : DocSoup<TableAnchor option> = 
    tableAreas >>>= fun arr -> 
    match Array.tryFindIndex (fun rgn -> isSubregionOf rgn needle) arr with
    | Some ix -> sreturn (Some <| TableAnchor ix)
    | None -> sreturn None


let containingTable (needle:Region) : DocSoup<TableAnchor> = 
    tableAreas >>>= fun arr -> 
    match Array.tryFindIndex (fun rgn -> isSubregionOf rgn needle) arr with
    | Some ix -> sreturn (TableAnchor ix)
    | None -> throwError "No containingTable found."


let containingCell (needle:Region) : DocSoup<CellAnchor> = 
    let rec work (tix:int) (tables:Word.Table list) : CellAnchor option = 
        let test (cell:Word.Cell) : bool = 
            isSubregionOf (extractRegion cell.Range) needle
        match tables with
        | [] -> None
        | (t :: ts) -> 
            match tryFindCell test t with
            | Some cell -> Some { TableIndex = TableAnchor tix;
                                    RowIndex = cell.RowIndex;
                                    ColumnIndex = cell.ColumnIndex }
            | None -> work (tix+1) ts

    getTablesInFocus >>>= fun arr -> 
    match work 0 (Array.toList arr) with
    | Some cid -> sreturn cid
    | None -> throwError "containingCell"

let private getTable (anchor:TableAnchor) : DocSoup<Word.Table> = 
    getTablesInFocus >>>= fun arr -> 
    try
        let ix = getTableIndex anchor
        sreturn arr.[ix]
    with
    | _ -> throwError "getTable error (index out-of-range?)"

let private getCell (anchor:CellAnchor) : DocSoup<Word.Cell> = 
    getTable (anchor.TableIndex) >>>= fun table ->
    try
        sreturn <| table.Cell(anchor.RowIndex, anchor.ColumnIndex)
    with
    | _ -> throwError "getCell error (index out-of-range?)"

        
let cellText (anchor:CellAnchor) : DocSoup<string> =
    getCell anchor >>>= fun whole -> focus (extractRegion whole.Range) getText

let tableText (anchor:TableAnchor) : DocSoup<string> =
    getTable anchor >>>= fun whole -> focus (extractRegion whole.Range) getText


let focusTable (anchor:TableAnchor) (ma:DocSoup<'a>) : DocSoup<'a> = 
    getTable anchor >>>= fun table ->
    let region = extractRegion table.Range 
    focus region ma


let focusCell (anchor:CellAnchor) (ma:DocSoup<'a>) : DocSoup<'a> = 
    getCell anchor >>>= fun cell ->
    let region = extractRegion cell.Range 
    focus region ma


let findCell (search:string) (matchCase:bool) : DocSoup<CellAnchor> =
    findTextMany search matchCase >>>= findSuccessM containingCell

let findCellPattern (search:string) : DocSoup<CellAnchor> =
    findPatternMany search >>>= findSuccessM containingCell

let findCells (search:string) (matchCase:bool) : DocSoup<CellAnchor list> =
    findTextMany search matchCase >>>= findSuccessesM containingCell


let findCellsPattern (search:string) : DocSoup<CellAnchor list> =
    findPatternMany search >>>= findSuccessesM containingCell



/// Finds first table containing search text.
/// If a match is found in "water" before a table, we continue the search.
let findTable (search:string) (matchCase:bool) : DocSoup<TableAnchor> =
    findTextMany search matchCase >>>= findSuccessM containingTable

/// Finds first table containing a match.
/// If a match is found in "water" before a table, we continue the search.
let findTablePattern (search:string) : DocSoup<TableAnchor> =
    findPatternMany search >>>= findSuccessM containingTable


/// Finds tables containing a match.
/// If a match is found in "water" before a table, we continue the search.    
let findTables (search:string) (matchCase:bool) : DocSoup<TableAnchor list> =
    findTextMany search matchCase >>>= findSuccessesM containingTable


let findTablesPattern (search:string) : DocSoup<TableAnchor list> =
    findPatternMany search >>>= findSuccessesM containingTable

/// Find first table that contains all the strings in the list of searches.
let findTableAll (searches:string list) (matchCase:bool) : DocSoup<TableAnchor> =
    let rec work (ss:string list) (anchors: TableAnchor list) = 
        match anchors with
        | [] -> throwError "findTableAll not found" 
        | (a1 :: rest) -> 
            focusTable a1 (optional (findAll ss matchCase)) >>>= fun ans ->
            match ans with
            | Some _ -> sreturn a1
            | None -> work ss rest
    match searches with
    | [] -> throwError "findTableAll empty search list"
    | [s] -> findTable s matchCase
    | (s :: ss) -> 
        findTables s matchCase >>>= fun tables -> 
        work ss tables

/// Find tables that contain all the strings in the list of searches.
let findTablesAll (searches:string list) (matchCase:bool) : DocSoup<TableAnchor list> =
    let rec work (ss:string list) (ac: TableAnchor list) (anchors: TableAnchor list)  = 
        match anchors with
        | [] -> sreturn (List.rev ac)
        | (a1 :: rest) -> 
            focusTable a1 (optional (findAll ss matchCase)) >>>= fun ans ->
            match ans with
            | Some _ -> work ss (a1::ac) rest
            | None -> work ss ac rest
    match searches with
    | [] -> throwError "findTablesAll - empty search list"
    | [s] -> findTables s matchCase
    | (s :: ss) -> 
        findTables s matchCase >>>= fun tables -> 
        work ss [] tables

let findTablePatternAll (searches:string list) : DocSoup<TableAnchor> =
    let rec work (ss:string list) (anchors: TableAnchor list) = 
        match anchors with
        | [] -> throwError "findTablePatternAll not found" 
        | (a1 :: rest) -> 
            focusTable a1 (optional (findPatternAll ss)) >>>= fun ans ->
            match ans with
            | Some _ -> sreturn a1
            | None -> work ss rest
    match searches with
    | [] -> throwError "findTablePatternAll empty search list"
    | [s] -> findTablePattern s
    | (s :: ss) -> 
        findTablesPattern s  >>>= fun tables -> 
        work ss tables

let findTablesPatternAll (searches:string list) : DocSoup<TableAnchor list> =
    let rec work (ss:string list) (ac: TableAnchor list) (anchors: TableAnchor list)  = 
        match anchors with
        | [] -> sreturn (List.rev ac)
        | (a1 :: rest) -> 
            focusTable a1 (optional (findPatternAll ss)) >>>= fun ans ->
            match ans with
            | Some _ -> work ss (a1::ac) rest
            | None -> work ss ac rest
    match searches with
    | [] -> throwError "findTablesAll - empty search list"
    | [s] -> findTablesPattern s 
    | (s :: ss) -> 
        findTablesPattern s >>>= fun tables -> 
        work ss [] tables

let assertInFocus (region:Region) : DocSoup<unit> = 
    DocSoup <| fun doc focus -> 
        if isSubregionOf focus region then
            Ok ()
        else
            Err "assertInFocus - outside focus"

let assertCellInFocus (anchor:CellAnchor) : DocSoup<unit> = 
    getCell anchor >>>= fun cell -> 
    assertInFocus (extractRegion cell.Range)
 
let assertTableInFocus (anchor:TableAnchor) : DocSoup<unit> = 
    getTable anchor >>>= fun table -> 
    assertInFocus (extractRegion table.Range)   


let parentTable (cell:CellAnchor) : DocSoup<TableAnchor> = 
    sreturn cell.TableIndex

let cellLeft (cell:CellAnchor) : DocSoup<CellAnchor> = 
    let c1 = { cell with ColumnIndex = cell.ColumnIndex - 1} 
    assertCellInFocus c1 >>>. sreturn c1


let cellRight (cell:CellAnchor) : DocSoup<CellAnchor> = 
    let c1 = { cell with ColumnIndex = cell.ColumnIndex + 1} 
    assertCellInFocus c1 >>>. sreturn c1

let cellBelow (cell:CellAnchor) : DocSoup<CellAnchor> = 
    let c1 = { cell with RowIndex = cell.RowIndex + 1} 
    assertCellInFocus c1 >>>. sreturn c1

let cellAbove (cell:CellAnchor) : DocSoup<CellAnchor> = 
    let c1 = { cell with RowIndex = cell.RowIndex - 1} 
    assertCellInFocus c1 >>>. sreturn c1

