// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocSoup.DocExtractor

open System.Text.RegularExpressions
open Microsoft.Office.Interop
open FParsec

open DocSoup.Base
open DocSoup.TableExtractor

            
type Cursor = int


type Result<'a> = 
    | Err of string
    | Ok of Cursor * 'a

let private resultConcat (source:Result<'a> list) : Result<'a list> = 
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

/// Operator for fmap.
let (|>>>) (ma:DocExtractor<'a>) (fn:'a -> 'b) : DocExtractor<'b> = 
    fmapM fn ma

/// Flipped fmap.
let (<<<|) (fn:'a -> 'b) (ma:DocExtractor<'a>) : DocExtractor<'b> = 
    fmapM fn ma

// liftM (which is fmap)
let liftM (fn:'a -> 'x) (ma:DocExtractor<'a>) : DocExtractor<'x> = 
    fmapM fn ma

let liftM2 (fn:'a -> 'b -> 'x) 
            (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'x> = 
    docExtract { 
        let! a = ma
        let! b = mb
        return (fn a b)
    }

let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
            (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) : DocExtractor<'x> = 
    docExtract { 
        let! a = ma
        let! b = mb
        let! c = mc
        return (fn a b c)
    }

let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
            (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) : DocExtractor<'x> = 
    docExtract { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        return (fn a b c d)
    }


let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) 
            (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
            (me:DocExtractor<'e>) : DocExtractor<'x> = 
    docExtract { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        let! e = me
        return (fn a b c d e)
    }

let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
            (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
            (me:DocExtractor<'e>) (mf:DocExtractor<'f>) : DocExtractor<'x> = 
    docExtract { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        let! e = me
        let! f = mf
        return (fn a b c d e f)
    }

let tupleM2 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'a * 'b> = 
    liftM2 (fun a b -> (a,b)) ma mb

let tupleM3 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) : DocExtractor<'a * 'b * 'c> = 
    liftM3 (fun a b c -> (a,b,c)) ma mb mc

let tupleM4 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) : DocExtractor<'a * 'b * 'c * 'd> = 
    liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

let tupleM5 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
            (me:DocExtractor<'e>) : DocExtractor<'a * 'b * 'c * 'd * 'e> = 
    liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

let tupleM6 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
            (me:DocExtractor<'e>) (mf:DocExtractor<'f>) : DocExtractor<'a * 'b * 'c * 'd * 'e * 'f> = 
    liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf

let pipeM2 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (fn:'a -> 'b -> 'x) : DocExtractor<'x> = 
    liftM2 fn ma mb

let pipeM3 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) 
            (fn:'a -> 'b -> 'c -> 'x): DocExtractor<'x> = 
    liftM3 fn ma mb mc

let pipeM4 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
            (fn:'a -> 'b -> 'c -> 'd -> 'x) : DocExtractor<'x> = 
    liftM4 fn ma mb mc md

let pipeM5 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
            (me:DocExtractor<'e>) 
            (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x): DocExtractor<'x> = 
    liftM5 fn ma mb mc md me

let pipeM6 (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) 
            (mc:DocExtractor<'c>) (md:DocExtractor<'d>) 
            (me:DocExtractor<'e>) (mf:DocExtractor<'f>) 
            (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x): DocExtractor<'x> = 
    liftM6 fn ma mb mc md me mf

/// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
let alt (ma:DocExtractor<'a>) (mb:DocExtractor<'a>) : DocExtractor<'a> = 
    DocExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err _ -> apply1 mb doc pos
        | Ok (pos1,a) -> Ok (pos1,a)

let (<||>) (ma:DocExtractor<'a>) (mb:DocExtractor<'a>) : DocExtractor<'a> = alt ma mb


// Haskell Applicative's (<*>)
let apM (mf:DocExtractor<'a ->'b>) (ma:DocExtractor<'a>) : DocExtractor<'b> = 
    docExtract { 
        let! fn = mf
        let! a = ma
        return (fn a) 
    }

let (<**>) (ma:DocExtractor<'a -> 'b>) (mb:DocExtractor<'a>) : DocExtractor<'b> = 
    apM ma mb

let (<&&>) (fn:'a -> 'b) (ma:DocExtractor<'a>) :DocExtractor<'b> = 
    fmapM fn ma


/// Perform two actions in sequence. Ignore the results of the second action if both succeed.
let seqL (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'a> = 
    docExtract { 
        let! a = ma
        let! b = mb
        return a
    }

/// Perform two actions in sequence. Ignore the results of the first action if both succeed.
let seqR (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'b> = 
    docExtract { 
        let! a = ma
        let! b = mb
        return b
    }

let (.>>>) (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'a> = 
    seqL ma mb

let (>>>.) (ma:DocExtractor<'a>) (mb:DocExtractor<'b>) : DocExtractor<'b> = 
    seqR ma mb


let mapM (p: 'a -> DocExtractor<'b>) (source:'a list) : DocExtractor<'b list> = 
    DocExtractor <| fun doc pos0 -> 
        let rec work pos ac ys = 
            match ys with
            | [] -> Ok (pos, List.rev ac)
            | z :: zs -> 
                match apply1 (p z) doc pos with
                | Err msg -> Err msg
                | Ok (pos1,ans) -> work pos1 (ans::ac) zs
        work pos0  [] source

let forM (source:'a list) (p: 'a -> DocExtractor<'b>) : DocExtractor<'b list> = 
    mapM p source




/// The action is expected to return ``true`` or `false``- if it throws 
/// an error then the error is passed upwards.
let findM  (action: 'a -> DocExtractor<bool>) (source:'a list) : DocExtractor<'a> = 
    DocExtractor <| fun doc pos0 -> 
        let rec work pos ys = 
            match ys with
            | [] -> Err "findM - not found"
            | z :: zs -> 
                match apply1 (action z) doc pos with
                | Err msg -> Err msg
                | Ok (pos1,ans) -> if ans then Ok (pos1,z) else work pos1 zs
        work pos0 source

/// The action is expected to return ``true`` or `false``- if it throws 
/// an error then the error is passed upwards.
let tryFindM (action: 'a -> DocExtractor<bool>) 
                (source:'a list) : DocExtractor<'a option> = 
    DocExtractor <| fun doc pos0 -> 
        let rec work pos ys = 
            match ys with
            | [] -> Ok (pos0,None)
            | z :: zs -> 
                match apply1 (action z) doc pos with
                | Err msg -> Err msg
                | Ok (pos1,ans) -> if ans then Ok (pos1, Some z) else work pos1 zs
        work pos0 source

    
let optionToFailure (ma:DocExtractor<option<'a>>) 
                    (errMsg:string) : DocExtractor<'a> = 
    DocExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err msg -> Err msg
        | Ok (_,None) -> Err errMsg
        | Ok (pos1, Some a) -> Ok (pos1,a)


/// Optionally parses. When the parser fails return None and don't move the cursor position.
let optional (ma:DocExtractor<'a>) : DocExtractor<'a option> = 
    DocExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err _ -> Ok (pos,None)
        | Ok (pos1,a) -> Ok (pos1,Some a)


let optionalz (ma:DocExtractor<'a>) : DocExtractor<unit> = 
    DocExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err _ -> Ok (pos, ())
        | Ok (pos1,_) -> Ok (pos1, ())

/// Turn an operation into a boolean, when the action is success return true 
/// when it fails return false
let boolify (ma:DocExtractor<'a>) : DocExtractor<bool> = 
    DocExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err _ -> Ok (pos,false)
        | Ok (pos1,_) -> Ok (pos1,true)

// *************************************
// Parser combinators

/// End of document?
let eof : DocExtractor<unit> =
    DocExtractor <| fun doc pos ->
        if pos >= doc.Range().End then 
            Ok (pos, ())
        else
            Err "eof (not-at-end)"


/// Parses p without consuming input
let lookahead (ma:DocExtractor<'a>) : DocExtractor<'a> =
    DocExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err msg -> Err msg
        | Ok (_,a) -> Ok (pos,a)



let between (popen:DocExtractor<_>) (pclose:DocExtractor<_>) 
            (ma:DocExtractor<'a>) : DocExtractor<'a> =
    docExtract { 
        let! _ = popen
        let! ans = ma
        let! _ = pclose
        return ans 
    }


let many (ma:DocExtractor<'a>) : DocExtractor<'a list> = 
    DocExtractor <| fun doc pos0 ->
        let rec work pos ac = 
            match apply1 ma doc pos with
            | Err _ -> Ok (pos, List.rev ac)
            | Ok (pos1,a) -> work pos1 (a::ac)
        work pos0 []

let many1 (ma:DocExtractor<'a>) : DocExtractor<'a list> = 
    docExtract { 
        let! a1 = ma
        let! rest = many ma
        return (a1::rest) 
    } 

let skipMany (ma:DocExtractor<'a>) : DocExtractor<unit> = 
    many ma >>>= fun _ -> dreturn ()

let sepBy1 (ma:DocExtractor<'a>) 
            (sep:DocExtractor<_>) : DocExtractor<'a list> = 
    docExtract { 
        let! a1 = ma
        let! rest = many (sep >>>. ma) 
        return (a1::rest)
    }

let sepBy (ma:DocExtractor<'a>) 
            (sep:DocExtractor<_>) : DocExtractor<'a list> = 
    sepBy1 ma sep <||> dreturn []

let manyTill (ma:DocExtractor<'a>) 
                (terminator:DocExtractor<_>) : DocExtractor<'a list> = 
    DocExtractor <| fun doc pos0 ->
        let rec work pos ac = 
            match apply1 terminator doc pos with
            | Err msg -> 
                match apply1 ma doc pos with
                | Err msg -> Err msg
                | Ok (pos1,a) -> work pos1 (a::ac) 
            | Ok (pos1,_) -> Ok(pos1, List.rev ac)
        work pos0 []

let manyTill1 (ma:DocExtractor<'a>) 
                (terminator:DocExtractor<_>) : DocExtractor<'a list> = 
    liftM2 (fun a xs -> a::xs) ma (manyTill ma terminator)



// *************************************
// Run functions



let runOnFile (ma:DocExtractor<'a>) (fileName:string) : Result<'a> =
    if System.IO.File.Exists (fileName) then
        let app = new Word.ApplicationClass (Visible = false) :> Word.Application
        try 
            let doc = app.Documents.Open(FileName = ref (fileName :> obj))
            let ans = apply1 ma doc (doc.Range().Start)
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


let runOnFileE (ma:DocExtractor<'a>) (fileName:string) : 'a =
    match runOnFile ma fileName with
    | Err msg -> failwith msg
    | Ok (_,a) -> a



// *************************************
// Run tableExtractor

let withTable (anchor:TableAnchor) (ma:TableExtractor<'a>) : DocExtractor<'a> = 
    DocExtractor <| fun doc _ ->
        try 
            let table:Word.Table = doc.Tables.Item(anchor.Index)
            match runTableExtractor (ma:TableExtractor<'a>) table with
            | TErr msg -> Err msg
            | TOk a -> 
                let pos1 = table.Range.End + 1
                Ok (pos1,a)
        with
        | _ -> Err "withTable" 

        
let withTableM (anchorQuery:DocExtractor<TableAnchor>) (ma:TableExtractor<'a>) : DocExtractor<'a> = 
    anchorQuery >>>= fun a -> withTable a ma 

// Now we have a cursor we can have a nextTable function.
let askNextTable : DocExtractor<TableAnchor> = 
    DocExtractor <| fun doc pos ->
        try 
            let needle = { RegionStart = pos; RegionEnd = pos+1 }
            let rec work (ix:TableAnchor) = 
                if ix.TableIndex <= doc.Tables.Count then 
                    let table = doc.Tables.Item (ix.TableIndex)
                    let region = extractRegion table.Range
                    if region.RegionStart >= pos then 
                        Ok (pos,ix)
                    else work ix.Next
                else
                    Err "askNextTable (no next table)"
            work TableAnchor.First
        with
        | _ -> Err "askNextTable" 


/// Note - this is unguarded, use with care in many, many1 etc. 
let nextTable (ma:TableExtractor<'a>) : DocExtractor<'a> = 
    withTableM askNextTable ma

// *************************************
// String level parsing with FParsec

// TODO - FParsec will have to run in regions so that we have a
// stopping boundary.

// We expect string level parsers might fail. 
// Use this with caution or use execFParsecFallback.
//let execFParsec (parser:ParsecParser<'a>) : DocExtractor<'a> = 
//    DocExtractor <| fun doc pos ->
//        match dict.GetText focus doc with
//        | None -> Err "execFParsec - no input text"
//        | Some text -> 
//            let name = doc.Name  
//            match runParserOnString parser () name text with
//            | Success(ans,_,_) -> Ok ans
//            | Failure(msg,_,_) -> Err msg



// Returns fallback text if FParsec fails.
//let execFParsecFallback (parser:ParsecParser<'a>) : DocExtractor<FParsecFallback<'a>> = 
//    DocExtractor <| fun doc pos ->
//        match dict.GetText focus doc with
//        | None -> Ok <| FallbackText ""
//        | Some text -> 
//            let name = doc.Name  
//            match runParserOnString parser () name text with
//            | Success(ans,_,_) -> Ok <| FParsecOk ans
//            | Failure(msg,_,_) -> Ok <| FallbackText text


// *************************************
// Errors

let throwError (msg:string) : DocExtractor<'a> = 
    DocExtractor <| fun _ _ -> Err msg

let swapError (msg:string) (ma:DocExtractor<'a>) : DocExtractor<'a> = 
    DocExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err _ -> Err msg
        | Ok (pos1,a) -> Ok (pos1,a)

let (<&?>) (ma:DocExtractor<'a>) (msg:string) : DocExtractor<'a> = 
    swapError msg ma

let (<?&>) (msg:string) (ma:DocExtractor<'a>) : DocExtractor<'a> = 
    swapError msg ma


// *************************************
// Search text for "anchors"

[<Struct>]
type SearchAnchor = 
    private SearchAnchor of int
        member v.Position = match v with | SearchAnchor i -> i

let private startOfRegion (region:Region) : SearchAnchor = 
    SearchAnchor region.RegionStart

let private endOfRegion (region:Region) : SearchAnchor = 
    SearchAnchor region.RegionEnd

let withSearchAnchor (anchor:SearchAnchor) 
                        (ma:DocExtractor<'a>) : DocExtractor<'a> = 
    DocExtractor <| fun doc pos  -> 
        if anchor.Position > pos then   
           apply1 ma doc anchor.Position
        else
           Err "withSearchAnchor"

let withSearchAnchorM (anchorQuery:DocExtractor<SearchAnchor>) 
                        (ma:DocExtractor<'a>) : DocExtractor<'a> = 
    anchorQuery >>>= fun a -> withSearchAnchor a ma

let advanceM (anchorQuery:DocExtractor<SearchAnchor>) : DocExtractor<unit> = 
    withSearchAnchorM anchorQuery (dreturn ())


let findText (search:string) (matchCase:bool) : DocExtractor<Region> =
    DocExtractor <| fun doc pos  -> 
        let range =  getRangeToEnd pos doc 
        match boundedFind1 search matchCase extractRegion range with
        | Some region -> Ok (pos, region)
        | None -> Err <| sprintf "findText - '%s' not found" search

let findTextStart (search:string) (matchCase:bool) : DocExtractor<SearchAnchor> =
    findText search matchCase |>>> startOfRegion

let findTextEnd (search:string) (matchCase:bool) : DocExtractor<SearchAnchor> =
    findText search matchCase |>>> endOfRegion

/// Case sensitivity always appears to be true for Wildcard matches.
let findPattern (search:string) : DocExtractor<Region> =
    DocExtractor <| fun doc pos  -> 
        let range =  getRangeToEnd pos doc 
        match boundedFindPattern1 search extractRegion range with
        | Some region -> Ok (pos, region)
        | None -> Err <| sprintf "findPattern - '%s' not found" search
        
let findPatternStart (search:string)  : DocExtractor<SearchAnchor> =
    findPattern search |>>> startOfRegion

let findPatternEnd (search:string) : DocExtractor<SearchAnchor> =
    findPattern search |>>> endOfRegion

let findTextMany (search:string) (matchCase:bool) : DocExtractor<Region list> =
    DocExtractor <| fun doc pos  -> 
        let range =  getRangeToEnd pos doc 
        let finds = boundedFindMany search matchCase extractRegion range
        Ok (pos,finds)

/// Case sensitivity always appears to be true for Wildcard matches.
let findPatternMany (search:string) : DocExtractor<Region list> =
    DocExtractor <| fun doc pos -> 
        let range =  getRangeToEnd pos doc 
        let finds = boundedFindPatternMany search extractRegion range
        Ok (pos,finds)


let whiteSpace : DocExtractor<unit> = 
    DocExtractor <| fun doc pos -> 
        match apply1 (findText "^w" false) doc pos with
        | Err _ -> Ok (pos, ())
        | Ok (_,region) -> 
            // findText can find a subsequent region...
            if region.RegionStart = pos then
                Ok (region.RegionEnd, ())
            else
                Ok (pos, ())


let whiteSpace1 : DocExtractor<unit> = 
    DocExtractor <| fun doc pos -> 
        match apply1 (findText "^w" false) doc pos with
        | Err _ -> Err "whiteSpace1"
        | Ok (_,region) -> 
            // findText can find a subsequent region... 
            if region.RegionStart = pos && region.RegionEnd > pos then 
                Ok (region.RegionEnd, ())
            else
                Err "whiteSpace1 (none)"
            

let private parseStringInternal (str:string) 
                                (matchCase:bool) : DocExtractor<string> = 
    DocExtractor <| fun doc pos -> 
        match apply1 (findText str matchCase) doc pos with
        | Err _ -> printfn "no find" ; Err "parseStringInternal"
        | Ok (_,region) -> 
            if region.RegionStart = pos then 
                let text = regionText region doc
                Ok (region.RegionEnd, text)
            else
                printfn "Pos=%i; RegionStart=%i; RegionEnd=%i" pos region.RegionStart region.RegionEnd
                Err "parseStringInternal"

let pstring (str:string) : DocExtractor<string> = 
    parseStringInternal str true <&?> "pstring"

let pstringCI (str:string) : DocExtractor<string> = 
    parseStringInternal str false <&?> "pstring"

let pchar (ch:char) : DocExtractor<char> = 
    docExtract { 
        let! s1 = parseStringInternal (ch.ToString()) true
        if s1.Length = 1 then 
            return s1.[0]
        else
            throwError "pchar" |> ignore
    }



let anyChar : DocExtractor<char> = 
    DocExtractor <| fun doc pos0 -> 
        try 
            let rec work pos = 
                let range = doc.Range(rbox pos, rbox <| pos+1)
                if Regex.Match(range.Text, @"\p{C}").Success then 
                    work (pos+1)
                else
                    let ch1 = cleanRangeText(range).[0]
                    Ok (pos+1, ch1)
            work pos0
        with
        | _ -> Err "anyChar"

