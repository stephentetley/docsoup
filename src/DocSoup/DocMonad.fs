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

let tupleM2 (ma:DocSoup<'a>) (mb:DocSoup<'b>) : DocSoup<'a * 'b> = 
    liftM2 (fun a b -> (a,b)) ma mb

let tupleM3 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) : DocSoup<'a * 'b * 'c> = 
    liftM3 (fun a b c -> (a,b,c)) ma mb mc

let tupleM4 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (md:DocSoup<'d>) : DocSoup<'a * 'b * 'c * 'd> = 
    liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

let tupleM5 (ma:DocSoup<'a>) (mb:DocSoup<'b>) (mc:DocSoup<'c>) (md:DocSoup<'d>) (me:DocSoup<'e>) : DocSoup<'a * 'b * 'c * 'd * 'e> = 
    liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

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

let (<??>) (ma:DocSoup<'a>) (msg:string) : DocSoup<'a> = swapError msg ma


let focus (region:Region) (ma:DocSoup<'a>) : DocSoup<'a> = 
    DocSoup <| fun doc _ -> apply1 ma doc region

let fparse (p:TextParser<'a>) : DocSoup<'a> = 
    DocSoup <| fun doc focus -> execFParsec doc focus p 


/// This gets the text within the cuurent focus!        
let getText : DocSoup<string> =
    DocSoup <| fun doc focus -> 
        let text = regionText focus doc 
        Ok text
        

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
            let rec work (ac:Word.Table list) (source:Word.Table list) = 
                match source with
                | [] -> List.rev ac |> List.toArray
                | (t :: ts) -> 
                    let tregion = t.Range |> extractRegion
                    if isSubregionOf focus tregion then 
                        work (t::ac) ts
                    else
                        work ac ts
            let wholeDoc:Word.Range = doc.Range()
            let allTables : Word.Table list = 
                doc.Range().Tables |> Seq.cast<Word.Table> |> Seq.toList
            Ok <| work [] allTables
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

let parentTable (cell:CellAnchor) : DocSoup<TableAnchor> = 
    sreturn cell.TableIndex

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


let focusTable (anchor:TableAnchor) (ma:DocSoup<'a>) : DocSoup<'a> = 
    getTable anchor >>>= fun table ->
    let region = extractRegion table.Range 
    focus region ma


let focusCell (anchor:CellAnchor) (ma:DocSoup<'a>) : DocSoup<'a> = 
    getCell anchor >>>= fun cell ->
    let region = extractRegion cell.Range 
    focus region ma


let findCell (search:string) (matchCase:bool) : DocSoup<CellAnchor> =
    findText search matchCase >>>= containingCell

let findCellPattern (search:string) : DocSoup<CellAnchor> =
    findPattern search >>>= containingCell

let findCells (search:string) (matchCase:bool) : DocSoup<CellAnchor list> =
    findTextMany search matchCase >>>= fun regions -> 
    mapM containingCell regions


let findCellsPattern (search:string) : DocSoup<CellAnchor list> =
    findPatternMany search >>>= fun regions -> 
    mapM containingCell regions


let findTable (search:string) (matchCase:bool) : DocSoup<TableAnchor> =
    findText search matchCase >>>= containingTable

let findTablePattern (search:string) : DocSoup<TableAnchor> =
    findPattern search >>>= containingTable
    
let findTables (search:string) (matchCase:bool) : DocSoup<TableAnchor list> =
    findTextMany search matchCase >>>= fun regions -> 
    mapM containingTable regions


let findTablesPattern (search:string) : DocSoup<TableAnchor list> =
    findPatternMany search >>>= fun regions -> 
    mapM containingTable regions


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