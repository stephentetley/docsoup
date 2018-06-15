﻿// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocSoup.DocExtractor

open Microsoft.Office.Interop
open FParsec

open DocSoup.Base
open DocSoup.TableExtractor

            
type Cursor = int


type Result<'a> = 
    | Err of string
    | Ok of Cursor * 'a

let resultConcat (source:Result<'a> list) : Result<'a list> = 
    let rec work pos ac xs = 
        match xs with
        | [] -> Ok (pos,List.rev ac)
        | Ok (pos1,a) :: ys -> work (max pos pos1) (a::ac) ys
        | Err msg :: _ -> Err msg
    work 1 [] source


// DocExtractor is Reader(immutable)+State+Error
type DocExtractor<'a> = 
    DocExtractor of (Word.Document -> Cursor -> Result<'a>)



let inline private apply1 (ma: DocExtractor<'a>) 
                            (doc: Word.Document) 
                            (pos: Cursor) : Result<'a>= 
    let (DocExtractor f) = ma in f doc pos

let inline dreturn (x:'a) : DocExtractor<'a> = 
    DocExtractor <| fun _ pos -> Ok (pos, x)


let inline private bindM (ma:DocExtractor<'a>) 
                            (f :'a -> DocExtractor<'b>) : DocExtractor<'b> =
    DocExtractor <| fun doc pos -> 
        match apply1 ma doc pos with
        | Err msg -> Err msg
        | Ok (pos1,a) -> apply1 (f a) doc pos1

let inline dzero () : DocExtractor<'a> = 
    DocExtractor <| fun _ _ -> Err "dzero"


let inline private combineM (ma:DocExtractor<unit>) 
                                (mb:DocExtractor<unit>) : DocExtractor<unit> = 
    DocExtractor <| fun doc pos -> 
        match apply1 ma doc pos with
        | Err msg -> Err msg
        | Ok (pos1,a) -> apply1 mb doc pos1


let inline private  delayM (fn:unit -> DocExtractor<'a>) : DocExtractor<'a> = 
    bindM (dreturn ()) fn 




type DocExtractorBuilder() = 
    member self.Return x            = dreturn x
    member self.Bind (p,f)          = bindM p f
    member self.Zero ()             = dzero ()
    member self.Combine (ma,mb)     = combineM ma mb
    member self.Delay fn            = delayM fn

// Prefer "parse" to "parser" for the _Builder instance

let (docExtract:DocExtractorBuilder) = new DocExtractorBuilder()


/// Bind operator (name avoids clash with FParsec).
let (>>>=) (ma:DocExtractor<'a>) 
            (fn:'a -> DocExtractor<'b>) : DocExtractor<'b> = 
    bindM ma fn


// Common monadic operations
let fmapM (fn:'a -> 'b) (ma:DocExtractor<'a>) : DocExtractor<'b> = 
    DocExtractor <| fun doc pos -> 
       match apply1 ma doc pos with
       | Err msg -> Err msg
       | Ok (pos1,a) -> Ok (pos1, fn a)

// This is the nub of embedding FParsec - name clashes.
// We avoid them by using longer names in DocSoup.
let (|>>>) (ma:DocExtractor<'a>) (fn:'a -> 'b) : DocExtractor<'b> = fmapM fn ma
let (<<<|) (fn:'a -> 'b) (ma:DocExtractor<'a>) : DocExtractor<'b> = fmapM fn ma

//let liftM (fn:'a -> 'x) (ma:DocExtractor<'a>) : DocExtractor<'x> = fmapM fn ma

//let liftM2 (fn:'a -> 'b -> 'x) 
//            (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'x> = 
//    docSoup { 
//        let! a = ma
//        let! b = mb
//        return (fn a b)
//    }

//let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
//            (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) : DocExtractor<'x> = 
//    docSoup { 
//        let! a = ma
//        let! b = mb
//        let! c = mc
//        return (fn a b c)
//    }

//let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
//            (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) : DocExtractor<'x> = 
//    docSoup { 
//        let! a = ma
//        let! b = mb
//        let! c = mc
//        let! d = md
//        return (fn a b c d)
//    }


//let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) 
//            (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
//            (me:DocExtractor<'e>) : DocExtractor<'x> = 
//    docSoup { 
//        let! a = ma
//        let! b = mb
//        let! c = mc
//        let! d = md
//        let! e = me
//        return (fn a b c d e)
//    }

//let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
//            (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
//            (me:DocExtractor<'e>) (mf:DocExtractor<'f>) : DocExtractor<'x> = 
//    docSoup { 
//        let! a = ma
//        let! b = mb
//        let! c = mc
//        let! d = md
//        let! e = me
//        let! f = mf
//        return (fn a b c d e f)
//    }

//let tupleM2 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'a * 'b> = 
//    liftM2 (fun a b -> (a,b)) ma mb

//let tupleM3 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) : DocExtractor<'a * 'b * 'c> = 
//    liftM3 (fun a b c -> (a,b,c)) ma mb mc

//let tupleM4 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) : DocExtractor<'a * 'b * 'c * 'd> = 
//    liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

//let tupleM5 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
//            (me:DocExtractor<'e>) : DocExtractor<'a * 'b * 'c * 'd * 'e> = 
//    liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

//let tupleM6 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
//            (me:DocExtractor<'e>) (mf:DocExtractor<'f>) : DocExtractor<'a * 'b * 'c * 'd * 'e * 'f> = 
//    liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf

//let pipeM2 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (fn:'a -> 'b -> 'x) : DocExtractor<'x> = 
//    liftM2 fn ma mb

//let pipeM3 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) 
//            (fn:'a -> 'b -> 'c -> 'x): DocExtractor<'x> = 
//    liftM3 fn ma mb mc

//let pipeM4 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
//            (fn:'a -> 'b -> 'c -> 'd -> 'x) : DocExtractor<'x> = 
//    liftM4 fn ma mb mc md

//let pipeM5 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
//            (me:DocExtractor<'e>) 
//            (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x): DocExtractor<'x> = 
//    liftM5 fn ma mb mc md me

//let pipeM6 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
//            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
//            (me:DocExtractor<'e>) (mf:DocExtractor<'f>) 
//            (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x): DocExtractor<'x> = 
//    liftM6 fn ma mb mc md me mf

// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
//let alt (ma:DocExtractor<'a>) (mb:DocExtractor<'a>) : DocExtractor<'a> = 
//    DocExtractor <| fun doc dict focus ->
//        match apply1 ma doc dict focus with
//        | Err _ -> apply1 mb doc dict focus
//        | Ok a -> Ok a

//let (<||>) (ma:DocExtractor<'a>) (mb:DocExtractor<'a>) : DocExtractor<'a> = alt ma mb


// Haskell Applicative's (<*>)
//let apM (mf:DocExtractor<'a ->'b>) (ma:DocExtractor<'a>) : DocExtractor<'b> = 
//    docSoup { 
//        let! fn = mf
//        let! a = ma
//        return (fn a) 
//    }

//let (<**>) (ma:DocExtractor<'a -> 'b>) (mb:DocExtractor<'a>) : DocExtractor<'b> = 
//    apM ma mb

//let (<&&>) (fn:'a -> 'b) (ma:DocExtractor<'a>) :DocExtractor<'b> = 
//    fmapM fn ma


// Perform two actions in sequence. Ignore the results of the second action if both succeed.
//let seqL (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'a> = 
//    docSoup { 
//        let! a = ma
//        let! b = mb
//        return a
//    }

// Perform two actions in sequence. Ignore the results of the first action if both succeed.
//let seqR (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'b> = 
//    docSoup { 
//        let! a = ma
//        let! b = mb
//        return b
//    }

//let (.>>>) (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'a> = 
//    seqL ma mb

//let (>>>.) (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'b> = 
//    seqR ma mb


//let mapM (p: 'a -> DocExtractor<'b>) (source:'a list) : DocExtractor<'b list> = 
//    DocExtractor <| fun doc dict focus -> 
//        let rec work ac ys = 
//            match ys with
//            | [] -> Ok <| List.rev ac
//            | z :: zs -> 
//                match apply1 (p z) doc dict focus with
//                | Err msg -> Err msg
//                | Ok ans -> work (ans::ac) zs
//        work [] source

//let forM (source:'a list) (p: 'a -> DocExtractor<'b>) : DocExtractor<'b list> = 
//    mapM p source




/// The action is expected to return ``true`` or `false``- if it throws 
/// an error then the error is passed upwards.
//let findM  (action: 'a -> DocExtractor<bool>) (source:'a list) : DocExtractor<'a> = 
//    DocExtractor <| fun doc dict focus -> 
//        let rec work ys = 
//            match ys with
//            | [] -> Err "findM - not found"
//            | z :: zs -> 
//                match apply1 (action z) doc dict focus with
//                | Err msg -> Err msg
//                | Ok ans -> if ans then Ok z else work zs
//        work source

/// The action is expected to return ``true`` or `false``- if it throws 
/// an error then the error is passed upwards.
//let tryFindM  (action: 'a -> DocExtractor<bool>) 
//                (source:'a list) : DocExtractor<'a option> = 
//    DocExtractor <| fun doc dict focus -> 
//        let rec work ys = 
//            match ys with
//            | [] -> Ok None
//            | z :: zs -> 
//                match apply1 (action z) doc dict focus with
//                | Err msg -> Err msg
//                | Ok ans -> if ans then Ok (Some z) else work zs
//        work source

    
//let optionToFailure (ma:DocExtractor<option<'a>>) 
//                    (errMsg:string) : DocExtractor<'a> = 
//    DocExtractor <| fun doc dict focus ->
//        match apply1 ma doc dict focus with
//        | Err msg -> Err msg
//        | Ok None -> Err errMsg
//        | Ok (Some a) -> Ok a


/// Captures errors thrown by the parser 
//let optional (ma:DocExtractor<'a>) : DocExtractor<'a option> = 
//    DocExtractor <| fun doc dict focus ->
//        match apply1 ma doc dict focus with
//        | Err _ -> Ok None
//        | Ok a -> Ok <| Some a


//let optionalz (ma:DocExtractor<'a>) : DocExtractor<unit> = 
//    DocExtractor <| fun doc dict focus ->
//        match apply1 ma doc dict focus with
//        | Err _ -> Ok ()
//        | Ok _ -> Ok ()

/// Turn an operation into a boolean, when the action is success return true 
/// when it fails return false
//let boolify (ma:DocExtractor<'a>) : DocExtractor<bool> = 
//    DocExtractor <| fun doc dict focus ->
//        match apply1 ma doc dict focus with
//        | Err _ -> Ok false
//        | Ok _ -> Ok true





// *************************************
// Run functions



//let runOnFile (ma:DocExtractor<'a>) (fileName:string) : Result<'a> =
//    if System.IO.File.Exists (fileName) then
//        let app = new Word.ApplicationClass (Visible = false) :> Word.Application
//        try 
//            let doc = app.Documents.Open(FileName = ref (fileName :> obj))
//            let region1 = maxRegion doc
//            let ans = apply1 ma doc documentFocus region1
//            doc.Close(SaveChanges = rbox false)
//            app.Quit()
//            ans
//        with
//        | ex -> 
//            try 
//                app.Quit ()
//                Err ex.Message
//            with
//            | _ -> Err ex.Message
                
//    else 
//        Err <| sprintf "Cannot find file %s" fileName


//let runOnFileE (ma:DocExtractor<'a>) (fileName:string) : 'a =
//    match runOnFile ma fileName with
//    | Err msg -> failwith msg
//    | Ok a -> a



// *************************************
// String level parsing with FParsec



// We expect string level parsers might fail. 
// Use this with caution or use execFParsecFallback.
//let execFParsec (parser:ParsecParser<'a>) : DocExtractor<'a> = 
//    DocExtractor <| fun doc dict focus ->
//        match dict.GetText focus doc with
//        | None -> Err "execFParsec - no input text"
//        | Some text -> 
//            let name = doc.Name  
//            match runParserOnString parser () name text with
//            | Success(ans,_,_) -> Ok ans
//            | Failure(msg,_,_) -> Err msg



// Returns fallback text if FParsec fails.
//let execFParsecFallback (parser:ParsecParser<'a>) : DocExtractor<FParsecFallback<'a>> = 
//    DocExtractor <| fun doc dict focus ->
//        match dict.GetText focus doc with
//        | None -> Ok <| FallbackText ""
//        | Some text -> 
//            let name = doc.Name  
//            match runParserOnString parser () name text with
//            | Success(ans,_,_) -> Ok <| FParsecOk ans
//            | Failure(msg,_,_) -> Ok <| FallbackText text


// *************************************
// Errors

//let throwError (msg:string) : DocExtractor<'a> = 
//    DocExtractor <| fun _ _ _ -> Err msg

//let swapError (msg:string) (ma:DocExtractor<'a>) : DocExtractor<'a> = 
//    DocExtractor <| fun doc dict focus ->
//        match apply1 ma doc dict focus with
//        | Err _ -> Err msg
//        | Ok a -> Ok a

//let (<&?>) (ma:DocExtractor<'a>) (msg:string) : DocExtractor<'a> = 
//    swapError msg ma

//let (<?&>) (msg:string) (ma:DocExtractor<'a>) : DocExtractor<'a> = 
//    swapError msg ma


// *************************************
// Control focus

//let askDocFocus : DocExtractor<DocumentAnchor> = 
//    DocExtractor <| fun _ _ focus -> Ok focus

//let askTableFocus : TableExtractor<TableAnchor> = 
//    DocExtractor <| fun _ _ focus -> Ok focus

//let askCellFocus : CellExtractor<CellAnchor> = 
//    DocExtractor <| fun _ _ focus -> Ok focus


/// Restrict focus to a part of the input doc identified by region.
/// Focus type stays the same
//let focus (region:'focus) (ma:DocExtractor<'a>) : DocExtractor<'a> = 
//    DocExtractor <| fun doc dict _ -> apply1 ma doc dict region

/// Version of focus that binds the region returned from a query.
/// Focus type stays the same
//let focusM (regionQuery:DocExtractor<'focus>) (ma:DocExtractor<'a>) : DocExtractor<'a> = 
//    regionQuery >>>= fun region -> focus region ma

/// Assert the supplied region is within focus.
//let assertInFocus (position:'focus) : DocExtractor<unit> = 
//    DocExtractor <| fun doc dict focus -> 
//        match dict.GetRegion focus doc, dict.GetRegion position doc  with
//        | Some focus1, Some position1 -> 
//            if isSubregionOf focus1 position1 then
//                Ok ()
//            else
//                Err "assertInFocus - outside focus"
//        | _, _ -> Err "assertInFocus - failure" 
 


/// Implementation note - this uses Word's table index (which is 1-indexed, IIRC)
/// Note, the actual index value should never be exposed to client code.
//let private getTable (anchor:TableAnchor) : DocExtractor< Word.Table> = 
//    DocExtractor <| fun doc dict focus -> 
//        match dict.GetRegion focus doc with
//        | None -> Err "getTable failure"
//        | Some focus1 -> 
//            match getTable anchor doc with
//            | None -> Err "getTable error (index out-of-range?)"
//            | Some table -> 
//                if isSubregionOf focus1 (extractRegion table.Range) then 
//                    Ok table
//                else
//                    Err "getTable error (not in focus)"


/// Note - potentially not all of the table might be in focus.
//let private getCell (anchor:CellAnchor) : DocExtractor<Word.Cell> = 
//    DocExtractor <| fun doc dict focus ->
//        match dict.GetRegion focus doc with
//        | None -> Err "getTable failure"
//        | Some focus1 -> 
//            match getCell anchor doc with
//            | None -> Err "getCell error (index out-of-range?)"
//            | Some cell -> 
//                if isSubregionOf focus1 (extractRegion cell.Range) then 
//                    Ok cell
//                else
//                    Err "getCell error (not in focus)"


/// Restrict focus to a part of the input doc identified by the table anchor.
//let focusTable (anchor:TableAnchor) (ma:DocSoup<TableAnchor,'a>) : DocExtractor<'a> = 
//    DocExtractor <| fun doc _ _ -> apply1 ma doc tableFocus anchor

/// Version of focusTable that binds the anchor returned from a query.
//let focusTableM (tableQuery:DocExtractor<TableAnchor>) 
//                (ma:DocSoup<TableAnchor,'a>) : DocExtractor<'a> = 
//    tableQuery >>>= fun anchor -> focusTable anchor ma

/// Restrict focus to a part of the input doc identified by the cell anchor.
//let focusCell (anchor:CellAnchor) (ma:DocSoup<CellAnchor,'a>) : DocExtractor<'a> = 
//    DocExtractor <| fun doc _ _ -> apply1 ma doc cellFocus anchor


/// Version of focusCell that binds the anchor returned from a query.
//let focusCellM (cellQuery:DocExtractor<CellAnchor>) (ma:DocSoup<CellAnchor,'a>) : DocExtractor<'a> = 
//    cellQuery >>>= fun anchor -> focusCell anchor ma


/// Assert the supplied table anchor is in focus.
//let assertTableInFocus (anchor:TableAnchor) : DocExtractor<unit> = 
//    DocExtractor <| fun doc dict focus -> 
//        match dict.GetRegion focus doc, tableFocus.GetRegion anchor doc  with
//        | Some focus1, Some position1 -> 
//            if isSubregionOf focus1 position1 then
//                Ok ()
//            else
//                Err "assertTableInFocus - outside focus"
//        | _, _ -> Err "assertTableInFocus" 
    
/// Assert the supplied cell anchor is in focus.
//let assertCellInFocus (anchor:CellAnchor) : DocExtractor<unit> = 
//    DocExtractor <| fun doc dict focus -> 
//        match dict.GetRegion focus doc, cellFocus.GetRegion anchor doc  with
//        | Some focus1, Some position1 -> 
//            if isSubregionOf focus1 position1 then
//                Ok ()
//            else
//                Err "assertTableInFocus - outside focus"
//        | _, _ -> 
//            Err "assertTableInFocus" 

// *************************************
// Retrieve input

/// This gets the text within the current focus.     
/// [Value restriction without ()]
//let private getText () : DocExtractor<string> =
//    DocExtractor <| fun doc dict focus -> 
//        match dict.GetText focus doc with
//        | None -> Err "getText"
//        | Some text -> Ok <| text.Trim ()



/// This gets all the text from a document that is within the current focus.
//let getFocusedText : DocExtractor<string> = getText ()


/// This gets all the text for the table under current focus.
//let getTableText : TableExtractor<string> = getText ()


/// This gets the text for the cell under current focus.
//let getCellText : CellExtractor<string> = getText ()



// *************************************
// Search text for "anchors"

//let findText (search:string) (matchCase:bool) : DocExtractor<Region> =
//    DocExtractor <| fun doc dict focus  -> 
//        match getRange dict focus doc with
//        | None -> Err "findText fail"
//        | Some range -> 
//            match boundedFind1 search matchCase extractRegion range with
//            | Some region -> Ok region
//            | None -> Err <| sprintf "findText - '%s' not found" search

/// Case sensitivity always appears to be true for Wildcard matches.
//let findPattern (search:string) : DocExtractor<Region> =
//    DocExtractor <| fun doc dict focus  -> 
//        match getRange dict focus doc with
//        | None -> Err "findPattern fail"
//        | Some range ->
//            match boundedFindPattern1 search extractRegion range with
//            | Some region -> Ok region
//            | None -> Err <| sprintf "findPattern - '%s' not found" search
        

//let findTextMany (search:string) (matchCase:bool) : DocExtractor<Region list> =
//    DocExtractor <| fun doc dict focus  -> 
//        match getRange dict focus doc with
//        | None -> Err "findTextMany"
//        | Some range ->
//            Ok <| boundedFindMany search matchCase extractRegion range


/// Case sensitivity always appears to be true for Wildcard matches.
//let findPatternMany (search:string) : DocExtractor<Region list> =
//    DocExtractor <| fun doc dict focus  -> 
//        match getRange dict focus doc with
//        | None -> Err "findPatternMany"
//        | Some range ->
//            Ok <| boundedFindPatternMany search extractRegion range





/// Return the table containing needle.
//let containingTable (needle:Region) : DocExtractor<TableAnchor> = 
//    DocExtractor <| fun doc dict focus -> 
//        let rec work (ix:TableAnchor) = 
//            if ix.TableIndex <= doc.Tables.Count then 
//                let table = doc.Tables.Item (ix.TableIndex)
//                if isSubregionOf (extractRegion table.Range) needle then
//                    Ok ix
//                else work ix.Next
//            else
//                Err "containingTable - needle out of range"
//        match dict.GetRegion focus doc with
//        | None -> Err "containingTable"
//        | Some focus1 ->
//            if isSubregionOf focus1 needle then
//                work TableAnchor.First
//            else
//                Err "containingTable - needle not in focus"


/// Return the cell containing needle.
//let containingCell (needle:Region) : DocExtractor<CellAnchor> = 
//    let testCell (cell:Word.Cell) : bool = 
//            isSubregionOf (extractRegion cell.Range) needle
//    docSoup { 
//        let! tableAnchor = containingTable needle
//        let! table = getTable tableAnchor
//        match tryFindCell testCell table with 
//        | Some cell -> 
//            return { 
//                TableIx = tableAnchor;
//                CellIx = { RowIx = cell.RowIndex; ColumnIx = cell.ColumnIndex }
//            }
//        | None -> throwError "containingCell - no match" |> ignore
//        }

    
/// Get the tableHeader region.        
//let tableHeader (anchor:TableAnchor) : DocExtractor<Region> = 
//    getCell (firstCell anchor) |>>> fun (cell:Word.Cell) -> extractRegion cell.Range



// *************************************
// Navigation

/// Get the table by index - must be in focus.
/// Note - indexing is from 1.
//let getTableByIndex (ix:int) : DocExtractor<TableAnchor> = 
//    let anchor = { TableIndex =  ix }
//    assertTableInFocus anchor >>>. sreturn anchor


/// Get the cell by index - must be in focus.
/// Note - indexing is from 1.
//let getCellByIndex (cellIx:CellIndex) : TableExtractor<CellAnchor> = 
//    askTableFocus |>>> fun tix -> { TableIx = tix; CellIx = cellIx }

/// Get the table containing the supplied cell.
//let parentTable (cell:CellAnchor) : DocExtractor<TableAnchor> = 
//    (assertTableInFocus cell.TableAnchor >>>. sreturn cell.TableAnchor) <&?> "parentTable - failed"



/// Get the next table, will fail if next table is not in focus
//let nextTable (anchor:TableAnchor) : DocExtractor<TableAnchor> = 
//    (assertTableInFocus anchor.Next >>>. sreturn anchor.Next) <&?> "nextTable - failed" 


//let cellLeft (cell:CellAnchor) : TableExtractor<CellAnchor> = 
//    let c1 = { cell with CellIx = cell.CellIx.DecrCol } 
//    (assertCellInFocus c1 >>>. sreturn c1) <&?> "cellLeft - failed" 


//let cellRight (cell:CellAnchor) : TableExtractor<CellAnchor> = 
//    let c1 = { cell with CellIx = cell.CellIx.IncrCol } 
//    (assertCellInFocus c1 >>>. sreturn c1) <&?> "cellRight - failed" 

//let cellBelow (cell:CellAnchor) : TableExtractor<CellAnchor> = 
//    let c1 =  { cell with CellIx = cell.CellIx.IncrRow } 
//    (assertCellInFocus c1 >>>. sreturn c1) <&?> "cellBelow - failed" 

//let cellAbove (cell:CellAnchor) : TableExtractor<CellAnchor> = 
//    let c1 =  { cell with CellIx = cell.CellIx.DecrRow }  
//    (assertCellInFocus c1 >>>. sreturn c1) <&?> "cellAbove - failed" 


// *************************************
// Find tables and within tables

        
//type private Finder<'a> = Word.Table -> option<'a>

//let exactFinder (search:string) (matchCase:bool) : Finder<Region> = 
//    fun (table:Word.Table) -> boundedFind1 search matchCase extractRegion table.Range

    
//let patternFinder (search:string) : Finder<Region> = 
//    fun (table:Word.Table) -> boundedFindPattern1 search extractRegion table.Range

//type private FinderMany<'a> = Word.Table -> 'a list

//let exactFinderMany (search:string) (matchCase:bool) : FinderMany<Region> = 
//    fun (table:Word.Table) -> boundedFindMany search matchCase extractRegion table.Range
    
//let patternFinderMany (search:string) : FinderMany<Region> = 
//    fun (table:Word.Table) -> boundedFindPatternMany search extractRegion table.Range
    


//let private findTableSingle (finder:Finder<'a>) : DocExtractor<TableAnchor> =
//    DocExtractor <| fun doc dict focus -> 
//        let tcount = doc.Tables.Count
//        let rec work (ix:TableAnchor) : Result<TableAnchor> = 
//            if ix.Index > tcount then
//                Err "findTableSingle - not found"
//            else
//                 Rather than fail if not in focus, move next instead 
//                 otherwise fail would short-curcuit.
//                 Note this masks index failures, hence tcount above.
//                match apply1 (getTable ix) doc dict focus with
//                | Err msg -> work ix.Next
//                | Ok table -> 
//                    match finder table with
//                    | None -> work ix.Next
//                    | Some _ -> Ok ix
//        work TableAnchor.First


//let private findTableMultiple (finder:Finder<'a>) : DocExtractor<TableAnchor list> =
//    DocExtractor <| fun doc dict focus -> 
//        let tcount = doc.Tables.Count
//        let rec work (ix:TableAnchor) (ac: TableAnchor list) : Result<TableAnchor list> = 
//            if ix.Index > tcount then
//                Ok <| List.rev ac 
//            else
//                 Rather than fail if not in focus, move next instead 
//                 otherwise fail would short-curcuit.
//                 Note this masks index failures, hence tcount above.
//                match apply1 (getTable ix) doc dict focus with
//                | Err msg -> work ix.Next ac
//                | Ok table-> 
//                    match finder table with
//                    | None -> work ix.Next ac
//                    | Some _ -> work ix.Next (ix::ac)
//        work TableAnchor.First []

/// Find the first table containing the search text.
//let findTable (search:string) (matchCase:bool) : DocExtractor<TableAnchor> = 
//    findTableSingle (exactFinder search matchCase)

/// Find the first table where the search pattern matches.
//let findTableByPattern (search:string) : DocExtractor<TableAnchor> = 
//    findTableSingle (patternFinder search)

/// Find all tables containing the search text.
//let findTables (search:string) (matchCase:bool) : DocExtractor<TableAnchor list> = 
//    findTableMultiple (exactFinder search matchCase)

/// Find all tables where the search pattern matches.
//let findTablesByPattern (search:string) : DocExtractor<TableAnchor list> = 
//    findTableMultiple (patternFinder search)


//let findCellSingle (finder:Finder<Region>) : TableExtractor<CellAnchor> =
//    DocExtractor <| fun doc dict focus -> 
//        match apply1 (getTable focus) doc dict focus with
//        | Err msg -> Err msg
//        | Ok table -> 
//            match finder table with
//            | None -> Err "findCell"
//            | Some region -> apply1 (containingCell region) doc dict focus

//let findCellMultiple (finder:FinderMany<Region>) : TableExtractor<CellAnchor list> =
//    DocExtractor <| fun doc dict focus -> 
//        let getAnchor (region:Region) = 
//            apply1 (containingCell region) doc dict focus 
//        match apply1 (getTable focus) doc dict focus with
//        | Err msg -> Err msg
//        | Ok table -> resultConcat << List.map getAnchor <| finder table 


/// Find a cell in the focused table.
//let findCell (search:string) (matchCase:bool) : TableExtractor<CellAnchor> =
//    findCellSingle (exactFinder search matchCase)

/// Find a cell in the focused table by wildcard pattern.
//let findCellPattern (search:string) : TableExtractor<CellAnchor> =
//    findCellSingle (patternFinder search)

/// Find possibly multiple cells in the focused table.
//let findCells (search:string) (matchCase:bool) : TableExtractor<CellAnchor list> =
//    findCellMultiple (exactFinderMany search matchCase)

/// Find possibly multiple cells in the focused table by wildcard pattern.
//let findCellsPattern (search:string) : TableExtractor<CellAnchor list> =
//    findCellMultiple (patternFinderMany search)
