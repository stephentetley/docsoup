// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocSoup.TablesExtractor

open System.Text.RegularExpressions
open Microsoft.Office.Interop
open FParsec

open DocSoup.Base
open DocSoup.TableExtractor1

            
type TableIx = int


type Result<'a> = 
    | Err of string
    | Ok of TableIx * 'a

let private resultConcat (source:Result<'a> list) : Result<'a list> = 
    let rec work pos ac xs = 
        match xs with
        | [] -> Ok (pos,List.rev ac)
        | Ok (pos1,a) :: ys -> work (max pos pos1) (a::ac) ys
        | Err msg :: _ -> Err msg
    work 1 [] source


// TablesExtractor is Reader(immutable)+State+Error
type TablesExtractor<'a> = 
    TablesExtractor of (Word.Table [] -> TableIx -> Result<'a>)



let inline private apply1 (ma: TablesExtractor<'a>) 
                            (tables: Word.Table []) 
                            (pos: TableIx) : Result<'a>= 
    let (TablesExtractor f) = ma in f tables pos

let inline treturn (x:'a) : TablesExtractor<'a> = 
    TablesExtractor <| fun _ pos -> Ok (pos, x)


let inline private bindM (ma:TablesExtractor<'a>) 
                            (f :'a -> TablesExtractor<'b>) : TablesExtractor<'b> =
    TablesExtractor <| fun doc pos -> 
        match apply1 ma doc pos with
        | Err msg -> Err msg
        | Ok (pos1,a) -> apply1 (f a) doc pos1

let inline tzero () : TablesExtractor<'a> = 
    TablesExtractor <| fun _ _ -> Err "tzero"


let inline private combineM (ma:TablesExtractor<unit>) 
                                (mb:TablesExtractor<unit>) : TablesExtractor<unit> = 
    TablesExtractor <| fun doc pos -> 
        match apply1 ma doc pos with
        | Err msg -> Err msg
        | Ok (pos1,a) -> apply1 mb doc pos1


let inline private  delayM (fn:unit -> TablesExtractor<'a>) : TablesExtractor<'a> = 
    bindM (treturn ()) fn 




type TablesExtractorBuilder() = 
    member self.Return x            = treturn x
    member self.Bind (p,f)          = bindM p f
    member self.Zero ()             = tzero ()
    member self.Combine (ma,mb)     = combineM ma mb
    member self.Delay fn            = delayM fn

// Prefer "parse" to "parser" for the _Builder instance

let (tablesExtract:TablesExtractorBuilder) = new TablesExtractorBuilder()

// *************************************
// Errors

let throwError (msg:string) : TablesExtractor<'a> = 
    TablesExtractor <| fun _ _ -> Err msg

let swapError (msg:string) (ma:TablesExtractor<'a>) : TablesExtractor<'a> = 
    TablesExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err _ -> Err msg
        | Ok (pos1,a) -> Ok (pos1,a)

let (<&?>) (ma:TablesExtractor<'a>) (msg:string) : TablesExtractor<'a> = 
    swapError msg ma

let (<?&>) (msg:string) (ma:TablesExtractor<'a>) : TablesExtractor<'a> = 
    swapError msg ma



/// Bind operator (name avoids clash with FParsec).
let (>>>=) (ma:TablesExtractor<'a>) 
            (fn:'a -> TablesExtractor<'b>) : TablesExtractor<'b> = 
    bindM ma fn


// Common monadic operations
let fmapM (fn:'a -> 'b) (ma:TablesExtractor<'a>) : TablesExtractor<'b> = 
    TablesExtractor <| fun doc pos -> 
       match apply1 ma doc pos with
       | Err msg -> Err msg
       | Ok (pos1,a) -> Ok (pos1, fn a)

// This is the nub of embedding FParsec - name clashes.
// We avoid them by using longer names in DocSoup.

/// Operator for fmap.
let (|>>>) (ma:TablesExtractor<'a>) (fn:'a -> 'b) : TablesExtractor<'b> = 
    fmapM fn ma

/// Flipped fmap.
let (<<<|) (fn:'a -> 'b) (ma:TablesExtractor<'a>) : TablesExtractor<'b> = 
    fmapM fn ma

// liftM (which is fmap)
let liftM (fn:'a -> 'x) (ma:TablesExtractor<'a>) : TablesExtractor<'x> = 
    fmapM fn ma

let liftM2 (fn:'a -> 'b -> 'x) 
            (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) : TablesExtractor<'x> = 
    tablesExtract { 
        let! a = ma
        let! b = mb
        return (fn a b)
    }

let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
            (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) : TablesExtractor<'x> = 
    tablesExtract { 
        let! a = ma
        let! b = mb
        let! c = mc
        return (fn a b c)
    }

let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
            (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) (md:TablesExtractor<'d>) : TablesExtractor<'x> = 
    tablesExtract { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        return (fn a b c d)
    }


let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) 
            (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) (md:TablesExtractor<'d>) 
            (me:TablesExtractor<'e>) : TablesExtractor<'x> = 
    tablesExtract { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        let! e = me
        return (fn a b c d e)
    }

let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
            (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) (md:TablesExtractor<'d>) 
            (me:TablesExtractor<'e>) (mf:TablesExtractor<'f>) : TablesExtractor<'x> = 
    tablesExtract { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        let! e = me
        let! f = mf
        return (fn a b c d e f)
    }

let tupleM2 (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) : TablesExtractor<'a * 'b> = 
    liftM2 (fun a b -> (a,b)) ma mb

let tupleM3 (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) : TablesExtractor<'a * 'b * 'c> = 
    liftM3 (fun a b c -> (a,b,c)) ma mb mc

let tupleM4 (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) (md:TablesExtractor<'d>) : TablesExtractor<'a * 'b * 'c * 'd> = 
    liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

let tupleM5 (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) (md:TablesExtractor<'d>) 
            (me:TablesExtractor<'e>) : TablesExtractor<'a * 'b * 'c * 'd * 'e> = 
    liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

let tupleM6 (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) (md:TablesExtractor<'d>) 
            (me:TablesExtractor<'e>) (mf:TablesExtractor<'f>) : TablesExtractor<'a * 'b * 'c * 'd * 'e * 'f> = 
    liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf

let pipeM2 (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (fn:'a -> 'b -> 'x) : TablesExtractor<'x> = 
    liftM2 fn ma mb

let pipeM3 (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) 
            (fn:'a -> 'b -> 'c -> 'x): TablesExtractor<'x> = 
    liftM3 fn ma mb mc

let pipeM4 (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) (md:TablesExtractor<'d>) 
            (fn:'a -> 'b -> 'c -> 'd -> 'x) : TablesExtractor<'x> = 
    liftM4 fn ma mb mc md

let pipeM5 (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) (md:TablesExtractor<'d>) 
            (me:TablesExtractor<'e>) 
            (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x): TablesExtractor<'x> = 
    liftM5 fn ma mb mc md me

let pipeM6 (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) (md:TablesExtractor<'d>) 
            (me:TablesExtractor<'e>) (mf:TablesExtractor<'f>) 
            (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x): TablesExtractor<'x> = 
    liftM6 fn ma mb mc md me mf

/// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
let alt (ma:TablesExtractor<'a>) (mb:TablesExtractor<'a>) : TablesExtractor<'a> = 
    TablesExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err _ -> apply1 mb doc pos
        | Ok (pos1,a) -> Ok (pos1,a)

let (<||>) (ma:TablesExtractor<'a>) (mb:TablesExtractor<'a>) : TablesExtractor<'a> = 
    alt ma mb <&?> "(<||>)"


// Haskell Applicative's (<*>)
let apM (mf:TablesExtractor<'a ->'b>) (ma:TablesExtractor<'a>) : TablesExtractor<'b> = 
    tablesExtract { 
        let! fn = mf
        let! a = ma
        return (fn a) 
    }

let (<**>) (ma:TablesExtractor<'a -> 'b>) (mb:TablesExtractor<'a>) : TablesExtractor<'b> = 
    apM ma mb

let (<&&>) (fn:'a -> 'b) (ma:TablesExtractor<'a>) :TablesExtractor<'b> = 
    fmapM fn ma


/// Perform two actions in sequence. Ignore the results of the second action if both succeed.
let seqL (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) : TablesExtractor<'a> = 
    tablesExtract { 
        let! a = ma
        let! b = mb
        return a
    }

/// Perform two actions in sequence. Ignore the results of the first action if both succeed.
let seqR (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) : TablesExtractor<'b> = 
    tablesExtract { 
        let! a = ma
        let! b = mb
        return b
    }

let (.>>>) (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) : TablesExtractor<'a> = 
    seqL ma mb

let (>>>.) (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) : TablesExtractor<'b> = 
    seqR ma mb


let mapM (p: 'a -> TablesExtractor<'b>) (source:'a list) : TablesExtractor<'b list> = 
    TablesExtractor <| fun doc pos0 -> 
        let rec work pos ac ys = 
            match ys with
            | [] -> Ok (pos, List.rev ac)
            | z :: zs -> 
                match apply1 (p z) doc pos with
                | Err msg -> Err msg
                | Ok (pos1,ans) -> work pos1 (ans::ac) zs
        work pos0  [] source

let forM (source:'a list) (p: 'a -> TablesExtractor<'b>) : TablesExtractor<'b list> = 
    mapM p source




/// The action is expected to return ``true`` or `false``- if it throws 
/// an error then the error is passed upwards.
let findM  (action: 'a -> TablesExtractor<bool>) (source:'a list) : TablesExtractor<'a> = 
    TablesExtractor <| fun doc pos0 -> 
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
let tryFindM (action: 'a -> TablesExtractor<bool>) 
                (source:'a list) : TablesExtractor<'a option> = 
    TablesExtractor <| fun doc pos0 -> 
        let rec work pos ys = 
            match ys with
            | [] -> Ok (pos0,None)
            | z :: zs -> 
                match apply1 (action z) doc pos with
                | Err msg -> Err msg
                | Ok (pos1,ans) -> if ans then Ok (pos1, Some z) else work pos1 zs
        work pos0 source

    
let optionToFailure (ma:TablesExtractor<option<'a>>) 
                    (errMsg:string) : TablesExtractor<'a> = 
    TablesExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err msg -> Err msg
        | Ok (_,None) -> Err errMsg
        | Ok (pos1, Some a) -> Ok (pos1,a)





/// Optionally parses. When the parser fails return None and don't move the cursor position.
let optional (ma:TablesExtractor<'a>) : TablesExtractor<'a option> = 
    TablesExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err _ -> Ok (pos,None)
        | Ok (pos1,a) -> Ok (pos1,Some a)


let optionalz (ma:TablesExtractor<'a>) : TablesExtractor<unit> = 
    TablesExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err _ -> Ok (pos, ())
        | Ok (pos1,_) -> Ok (pos1, ())

/// Turn an operation into a boolean, when the action is success return true 
/// when it fails return false
let boolify (ma:TablesExtractor<'a>) : TablesExtractor<bool> = 
    TablesExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err _ -> Ok (pos,false)
        | Ok (pos1,_) -> Ok (pos1,true)

// *************************************
// Parser combinators

/// End of document?
let eof : TablesExtractor<unit> =
    TablesExtractor <| fun tables pos ->
        if pos >= tables.Length then 
            Ok (pos, ())
        else
            Err "eof (not-at-end)"


/// Parses p without consuming input
let lookahead (ma:TablesExtractor<'a>) : TablesExtractor<'a> =
    TablesExtractor <| fun doc pos ->
        match apply1 ma doc pos with
        | Err msg -> Err msg
        | Ok (_,a) -> Ok (pos,a)



let between (popen:TablesExtractor<_>) (pclose:TablesExtractor<_>) 
            (ma:TablesExtractor<'a>) : TablesExtractor<'a> =
    tablesExtract { 
        let! _ = popen
        let! ans = ma
        let! _ = pclose
        return ans 
    }


let many (ma:TablesExtractor<'a>) : TablesExtractor<'a list> = 
    TablesExtractor <| fun doc pos0 ->
        let rec work pos ac = 
            match apply1 ma doc pos with
            | Err _ -> Ok (pos, List.rev ac)
            | Ok (pos1,a) -> work pos1 (a::ac)
        work pos0 []

let many1 (ma:TablesExtractor<'a>) : TablesExtractor<'a list> = 
    tablesExtract { 
        let! a1 = ma
        let! rest = many ma
        return (a1::rest) 
    } 

let skipMany (ma:TablesExtractor<'a>) : TablesExtractor<unit> = 
    many ma >>>= fun _ -> treturn ()

let sepBy1 (ma:TablesExtractor<'a>) 
            (sep:TablesExtractor<_>) : TablesExtractor<'a list> = 
    tablesExtract { 
        let! a1 = ma
        let! rest = many (sep >>>. ma) 
        return (a1::rest)
    }

let sepBy (ma:TablesExtractor<'a>) 
            (sep:TablesExtractor<_>) : TablesExtractor<'a list> = 
    sepBy1 ma sep <||> treturn []

let manyTill (ma:TablesExtractor<'a>) 
                (terminator:TablesExtractor<_>) : TablesExtractor<'a list> = 
    TablesExtractor <| fun doc pos0 ->
        let rec work pos ac = 
            match apply1 terminator doc pos with
            | Err msg -> 
                match apply1 ma doc pos with
                | Err msg -> Err msg
                | Ok (pos1,a) -> work pos1 (a::ac) 
            | Ok (pos1,_) -> Ok(pos1, List.rev ac)
        work pos0 []

let manyTill1 (ma:TablesExtractor<'a>) 
                (terminator:TablesExtractor<_>) : TablesExtractor<'a list> = 
    liftM2 (fun a xs -> a::xs) ma (manyTill ma terminator)



// *************************************
// Run functions



let runOnFile (ma:TablesExtractor<'a>) (fileName:string) : Result<'a> =
    if System.IO.File.Exists (fileName) then
        let app = new Word.ApplicationClass (Visible = false) :> Word.Application
        try 
            let doc = app.Documents.Open(FileName = ref (fileName :> obj))
            let tables : Word.Table [] = 
                doc.Tables |> Seq.cast<Word.Table> |> Seq.toArray
            let ans = apply1 ma tables 0
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


let runOnFileE (ma:TablesExtractor<'a>) (fileName:string) : 'a =
    match runOnFile ma fileName with
    | Err msg -> failwith msg
    | Ok (_,a) -> a



//// *************************************
//// Run tableExtractor

//let withTable (anchor:TableAnchor) (ma:TablesExtractor1<'a>) : TablesExtractor<'a> = 
//    TablesExtractor <| fun doc _ ->
//        try 
//            let table:Word.Table = doc.Tables.Item(anchor.Index)
//            match runTablesExtractor1 (ma:TablesExtractor1<'a>) table with
//            | T1Err msg -> Err msg
//            | T1Ok a -> 
//                let pos1 = table.Range.End + 1
//                Ok (pos1,a)
//        with
//        | _ -> Err "withTable" 

        
//let withTableM (anchorQuery:TablesExtractor<TableAnchor>) (ma:TablesExtractor1<'a>) : TablesExtractor<'a> = 
//    anchorQuery >>>= fun a -> withTable a ma 

//// Now we have a cursor we can have a nextTable function.
//let askNextTable : TablesExtractor<TableAnchor> = 
//    TablesExtractor <| fun doc pos ->
//        try 
//            let needle = { RegionStart = pos; RegionEnd = pos+1 }
//            let rec work (ix:TableAnchor) = 
//                if ix.TableIndex <= doc.Tables.Count then 
//                    let table = doc.Tables.Item (ix.TableIndex)
//                    let region = extractRegion table.Range
//                    if region.RegionStart >= pos then 
//                        Ok (pos,ix)
//                    else work ix.Next
//                else
//                    Err "askNextTable (no next table)"
//            work TableAnchor.First
//        with
//        | _ -> Err "askNextTable" 

///// This is the wrong abstraction / traversal strategy
///// We should be using the index into the array of tables in a document.


///// Note - this is unguarded, use with care in many, many1 etc. 
//let nextTable (ma:TablesExtractor1<'a>) : TablesExtractor<'a> = 
//    withTableM askNextTable ma

//// *************************************
//// String level parsing with FParsec

//// TODO - FParsec will have to run in regions so that we have a
//// stopping boundary.

//// We expect string level parsers might fail. 
//// Use this with caution or use execFParsecFallback.
////let execFParsec (parser:ParsecParser<'a>) : TablesExtractor<'a> = 
////    TablesExtractor <| fun doc pos ->
////        match dict.GetText focus doc with
////        | None -> Err "execFParsec - no input text"
////        | Some text -> 
////            let name = doc.Name  
////            match runParserOnString parser () name text with
////            | Success(ans,_,_) -> Ok ans
////            | Failure(msg,_,_) -> Err msg



//// Returns fallback text if FParsec fails.
////let execFParsecFallback (parser:ParsecParser<'a>) : TablesExtractor<FParsecFallback<'a>> = 
////    TablesExtractor <| fun doc pos ->
////        match dict.GetText focus doc with
////        | None -> Ok <| FallbackText ""
////        | Some text -> 
////            let name = doc.Name  
////            match runParserOnString parser () name text with
////            | Success(ans,_,_) -> Ok <| FParsecOk ans
////            | Failure(msg,_,_) -> Ok <| FallbackText text




//// *************************************
//// Search text for "anchors"

//[<Struct>]
//type SearchAnchor = 
//    private SearchAnchor of int
//        member v.Position = match v with | SearchAnchor i -> i

//let private startOfRegion (region:Region) : SearchAnchor = 
//    SearchAnchor region.RegionStart

//let private endOfRegion (region:Region) : SearchAnchor = 
//    SearchAnchor region.RegionEnd

//let withSearchAnchor (anchor:SearchAnchor) 
//                        (ma:TablesExtractor<'a>) : TablesExtractor<'a> = 
//    TablesExtractor <| fun doc pos  -> 
//        if anchor.Position > pos then   
//           apply1 ma doc anchor.Position
//        else
//           Err "withSearchAnchor"

//let withSearchAnchorM (anchorQuery:TablesExtractor<SearchAnchor>) 
//                        (ma:TablesExtractor<'a>) : TablesExtractor<'a> = 
//    anchorQuery >>>= fun a -> withSearchAnchor a ma

//let advanceM (anchorQuery:TablesExtractor<SearchAnchor>) : TablesExtractor<unit> = 
//    withSearchAnchorM anchorQuery (dreturn ())


//let findText (search:string) (matchCase:bool) : TablesExtractor<Region> =
//    TablesExtractor <| fun doc pos  -> 
//        let range =  getRangeToEnd pos doc 
//        match boundedFind1 search matchCase extractRegion range with
//        | Some region -> Ok (pos, region)
//        | None -> Err <| sprintf "findText - '%s' not found" search

//let findTextStart (search:string) (matchCase:bool) : TablesExtractor<SearchAnchor> =
//    findText search matchCase |>>> startOfRegion

//let findTextEnd (search:string) (matchCase:bool) : TablesExtractor<SearchAnchor> =
//    findText search matchCase |>>> endOfRegion

///// Case sensitivity always appears to be true for Wildcard matches.
//let findPattern (search:string) : TablesExtractor<Region> =
//    TablesExtractor <| fun doc pos  -> 
//        let range =  getRangeToEnd pos doc 
//        match boundedFindPattern1 search extractRegion range with
//        | Some region -> Ok (pos, region)
//        | None -> Err <| sprintf "findPattern - '%s' not found" search
        
//let findPatternStart (search:string)  : TablesExtractor<SearchAnchor> =
//    findPattern search |>>> startOfRegion

//let findPatternEnd (search:string) : TablesExtractor<SearchAnchor> =
//    findPattern search |>>> endOfRegion

//let findTextMany (search:string) (matchCase:bool) : TablesExtractor<Region list> =
//    TablesExtractor <| fun doc pos  -> 
//        let range =  getRangeToEnd pos doc 
//        let finds = boundedFindMany search matchCase extractRegion range
//        Ok (pos,finds)

///// Case sensitivity always appears to be true for Wildcard matches.
//let findPatternMany (search:string) : TablesExtractor<Region list> =
//    TablesExtractor <| fun doc pos -> 
//        let range =  getRangeToEnd pos doc 
//        let finds = boundedFindPatternMany search extractRegion range
//        Ok (pos,finds)


//// *************************************
//// Character level parsers

//let whiteSpace : TablesExtractor<string> = 
//    TablesExtractor <| fun doc pos -> 
//        try 
//            let range = doc.Range(Start=rbox pos)
//            let regMatch = Regex.Match(range.Text, @"^[\p{C}-[\r\n]]+")
//            if regMatch.Success then 
//                let pos1 = pos + regMatch.Length + 1
//                Ok (pos1, regMatch.Value)
//            else 
//                Ok (pos,"")
//        with
//        | _ -> Err "whiteSpace"

//let whiteSpace1 : TablesExtractor<string> = 
//    TablesExtractor <| fun doc pos -> 
//        try 
//            let range = doc.Range(Start=rbox pos)
//            let regMatch = Regex.Match(range.Text, @"^[\p{C}-[\r\n]]+")
//            if regMatch.Success then 
//                let pos1 = pos + regMatch.Length + 1
//                Ok (pos1, regMatch.Value)
//            else 
//                Err "whiteSpace1"
//        with
//        | _ -> Err "whiteSpace1"


///// This ignores control characters.
//let internal removeControlPrefix : TablesExtractor<unit> = 
//    TablesExtractor <| fun doc pos -> 
//        try 
//            let range = doc.Range(Start=rbox pos)
//            let regMatch = Regex.Match(range.Text, @"\A[\p{C}]+")
//            if regMatch.Success then 
//                let pos1 = pos + regMatch.Length
//                Ok (pos1, ())
//            else
//                Ok (pos, ())
//        with
//        | _ -> Err "removeContrlPrefix"

//let private parseStringInternal (str:string) 
//                                (matchCase:bool) : TablesExtractor<string> = 
//    TablesExtractor <| fun doc pos -> 
//        match apply1 (findText str matchCase) doc pos with
//        | Err _ -> Err "parseStringInternal"
//        | Ok (_,region) -> 
//            if region.RegionStart = pos then 
//                let text = regionText region doc
//                Ok (region.RegionEnd, text)
//            else
//                Err "parseStringInternal"

//let pstring (str:string) : TablesExtractor<string> = 
//    removeControlPrefix >>>. parseStringInternal str true <&?> "pstring"

//let pstringCI (str:string) : TablesExtractor<string> = 
//    removeControlPrefix >>>. parseStringInternal str false <&?> "pstringCI"

//let pchar (ch:char) : TablesExtractor<char> = 
//    tablesExtract { 
//        let! s1 = pstring (ch.ToString()) 
//        if s1.Length = 1 then 
//            return s1.[0]
//        else
//            throwError "pchar" |> ignore
//    } <&?> "pchar"


///// This ignores control characters.
//let anyChar : TablesExtractor<char> = 
//    TablesExtractor <| fun doc pos -> 
//        try 
//            let range = doc.Range(Start=rbox pos)
//            let regMatch = Regex.Match(range.Text, @"\A[\p{C}]+")
//            if regMatch.Success then 
//                let pos1 = pos + regMatch.Length
//                let range1 = doc.Range(Start = rbox pos1)
//                let ch1 = cleanRangeText(range1).[0]
//                Ok (pos1+1, ch1)
//            else
//                let range1 = doc.Range(Start=rbox pos)
//                let ch1 = cleanRangeText(range1).[0]
//                Ok (pos+1, ch1)
//        with
//        | _ -> Err "anyChar"

