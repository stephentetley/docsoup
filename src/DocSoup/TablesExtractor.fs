// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocSoup.TablesExtractor

open System.Text.RegularExpressions
open Microsoft.Office.Interop
open FParsec

open DocSoup.Base
open DocSoup.RowExtractor
open DocSoup.TableExtractor1

            
type TableIx = int


type Result<'a> = 
    | Err of string
    | Ok of TableIx * 'a

let private resultConcat (source:Result<'a> list) : Result<'a list> = 
    let rec work ix ac xs = 
        match xs with
        | [] -> Ok (ix,List.rev ac)
        | Ok (ix1,a) :: ys -> work (max ix ix1) (a::ac) ys
        | Err msg :: _ -> Err msg
    work 1 [] source


// TablesExtractor is Reader(immutable)+State+Error
type TablesExtractor<'a> = 
    TablesExtractor of (Word.Table [] -> TableIx -> Result<'a>)



let inline private apply1 (ma: TablesExtractor<'a>) 
                            (tables: Word.Table []) 
                            (ix: TableIx) : Result<'a>= 
    let (TablesExtractor f) = ma in f tables ix

let inline treturn (x:'a) : TablesExtractor<'a> = 
    TablesExtractor <| fun _ ix -> Ok (ix, x)


let inline private bindM (ma:TablesExtractor<'a>) 
                            (f :'a -> TablesExtractor<'b>) : TablesExtractor<'b> =
    TablesExtractor <| fun tables ix -> 
        match apply1 ma tables ix with
        | Err msg -> Err msg
        | Ok (ix1,a) -> apply1 (f a) tables ix1

let inline tzero () : TablesExtractor<'a> = 
    TablesExtractor <| fun _ _ -> Err "tzero"


let inline private combineM (ma:TablesExtractor<unit>) 
                                (mb:TablesExtractor<unit>) : TablesExtractor<unit> = 
    TablesExtractor <| fun tables ix -> 
        match apply1 ma tables ix with
        | Err msg -> Err msg
        | Ok (ix1,a) -> apply1 mb tables ix1


let inline private  delayM (fn:unit -> TablesExtractor<'a>) : TablesExtractor<'a> = 
    bindM (treturn ()) fn 




type TablesExtractorBuilder() = 
    member self.Return x            = treturn x
    member self.Bind (p,f)          = bindM p f
    member self.Zero ()             = tzero ()
    member self.Combine (ma,mb)     = combineM ma mb
    member self.Delay fn            = delayM fn

// Prefer "parse" to "parser" for the _Builder instance

let (parseTables:TablesExtractorBuilder) = new TablesExtractorBuilder()

// *************************************
// Errors

let throwError (msg:string) : TablesExtractor<'a> = 
    TablesExtractor <| fun _ _ -> Err msg

let swapError (msg:string) (ma:TablesExtractor<'a>) : TablesExtractor<'a> = 
    TablesExtractor <| fun tables ix ->
        match apply1 ma tables ix with
        | Err _ -> Err msg
        | Ok (ix1,a) -> Ok (ix1,a)

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
    TablesExtractor <| fun tables ix -> 
       match apply1 ma tables ix with
       | Err msg -> Err msg
       | Ok (ix1,a) -> Ok (ix1, fn a)

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
    parseTables { 
        let! a = ma
        let! b = mb
        return (fn a b)
    }

let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
            (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) : TablesExtractor<'x> = 
    parseTables { 
        let! a = ma
        let! b = mb
        let! c = mc
        return (fn a b c)
    }

let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
            (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) 
            (mc:TablesExtractor<'c>) (md:TablesExtractor<'d>) : TablesExtractor<'x> = 
    parseTables { 
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
    parseTables { 
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
    parseTables { 
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
    TablesExtractor <| fun tables ix ->
        match apply1 ma tables ix with
        | Err _ -> apply1 mb tables ix
        | Ok (ix1,a) -> Ok (ix1,a)

let (<||>) (ma:TablesExtractor<'a>) (mb:TablesExtractor<'a>) : TablesExtractor<'a> = 
    alt ma mb <&?> "(<||>)"


// Haskell Applicative's (<*>)
let apM (mf:TablesExtractor<'a ->'b>) (ma:TablesExtractor<'a>) : TablesExtractor<'b> = 
    parseTables { 
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
    parseTables { 
        let! a = ma
        let! b = mb
        return a
    }

/// Perform two actions in sequence. Ignore the results of the first action if both succeed.
let seqR (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) : TablesExtractor<'b> = 
    parseTables { 
        let! a = ma
        let! b = mb
        return b
    }

let (.>>>) (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) : TablesExtractor<'a> = 
    seqL ma mb

let (>>>.) (ma:TablesExtractor<'a>) (mb:TablesExtractor<'b>) : TablesExtractor<'b> = 
    seqR ma mb


let mapM (p: 'a -> TablesExtractor<'b>) (source:'a list) : TablesExtractor<'b list> = 
    TablesExtractor <| fun tables ix0 -> 
        let rec work ix ac ys = 
            match ys with
            | [] -> Ok (ix, List.rev ac)
            | z :: zs -> 
                match apply1 (p z) tables ix with
                | Err msg -> Err msg
                | Ok (ix1,ans) -> work ix1 (ans::ac) zs
        work ix0  [] source

let forM (source:'a list) (p: 'a -> TablesExtractor<'b>) : TablesExtractor<'b list> = 
    mapM p source

    
let optionToFailure (ma:TablesExtractor<option<'a>>) 
                    (errMsg:string) : TablesExtractor<'a> = 
    TablesExtractor <| fun tables ix ->
        match apply1 ma tables ix with
        | Err msg -> Err msg
        | Ok (_,None) -> Err errMsg
        | Ok (ix1, Some a) -> Ok (ix1,a)





/// Optionally parses. When the parser fails return None and don't move the cursor position.
let optional (ma:TablesExtractor<'a>) : TablesExtractor<'a option> = 
    TablesExtractor <| fun tables ix ->
        match apply1 ma tables ix with
        | Err _ -> Ok (ix,None)
        | Ok (ix1,a) -> Ok (ix1,Some a)


let optionalz (ma:TablesExtractor<'a>) : TablesExtractor<unit> = 
    TablesExtractor <| fun tables ix ->
        match apply1 ma tables ix with
        | Err _ -> Ok (ix, ())
        | Ok (ix1,_) -> Ok (ix1, ())

/// Turn an operation into a boolean, when the action is success return true 
/// when it fails return false
let boolify (ma:TablesExtractor<'a>) : TablesExtractor<bool> = 
    TablesExtractor <| fun tables ix ->
        match apply1 ma tables ix with
        | Err _ -> Ok (ix,false)
        | Ok (ix1,_) -> Ok (ix1,true)

// *************************************
// Parser combinators

/// End of document?
let eof : TablesExtractor<unit> =
    TablesExtractor <| fun tables ix ->
        if ix >= tables.Length then 
            Ok (ix, ())
        else
            Err "eof (not-at-end)"


/// Parses p without consuming input
let lookahead (ma:TablesExtractor<'a>) : TablesExtractor<'a> =
    TablesExtractor <| fun tables ix ->
        match apply1 ma tables ix with
        | Err msg -> Err msg
        | Ok (_,a) -> Ok (ix,a)



let between (popen:TablesExtractor<_>) (pclose:TablesExtractor<_>) 
            (ma:TablesExtractor<'a>) : TablesExtractor<'a> =
    parseTables { 
        let! _ = popen
        let! ans = ma
        let! _ = pclose
        return ans 
    }


let many (ma:TablesExtractor<'a>) : TablesExtractor<'a list> = 
    TablesExtractor <| fun tables ix0 ->
        let rec work ix ac = 
            match apply1 ma tables ix with
            | Err _ -> Ok (ix, List.rev ac)
            | Ok (ix1,a) -> work ix1 (a::ac)
        work ix0 []

let many1 (ma:TablesExtractor<'a>) : TablesExtractor<'a list> = 
    parseTables { 
        let! a1 = ma
        let! rest = many ma
        return (a1::rest) 
    } 

let skipMany (ma:TablesExtractor<'a>) : TablesExtractor<unit> = 
    many ma >>>= fun _ -> treturn ()

let sepBy1 (ma:TablesExtractor<'a>) 
            (sep:TablesExtractor<_>) : TablesExtractor<'a list> = 
    parseTables { 
        let! a1 = ma
        let! rest = many (sep >>>. ma) 
        return (a1::rest)
    }

let sepBy (ma:TablesExtractor<'a>) 
            (sep:TablesExtractor<_>) : TablesExtractor<'a list> = 
    sepBy1 ma sep <||> treturn []

let manyTill (ma:TablesExtractor<'a>) 
                (terminator:TablesExtractor<_>) : TablesExtractor<'a list> = 
    TablesExtractor <| fun tables ix0 ->
        let rec work ix ac = 
            match apply1 terminator tables ix with
            | Err msg -> 
                match apply1 ma tables ix with
                | Err msg -> Err msg
                | Ok (ix1,a) -> work ix1 (a::ac) 
            | Ok (ix1,_) -> Ok(ix1, List.rev ac)
        work ix0 []

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
//// Run tableExtractor1

let parseTable (ma:Table1<'a>) : TablesExtractor<'a> = 
    TablesExtractor <| fun tables ix ->
        try 
            let table:Word.Table = tables.[ix]
            match runTable1 ma table with
            | T1Err msg -> Err msg
            | T1Ok a -> Ok (ix+1,a)
        with
        | _ -> Err "parseTable" 

        

//// *************************************
//// Run RowExtractor

let parseTableRowwise (ma:RowExtractor<'a>) : TablesExtractor<'a> = 
    TablesExtractor <| fun tables ix ->
        try 
            let table:Word.Table = tables.[ix]
            match runRowExtractor ma table with
            | RErr msg -> Err msg
            | ROk (_,a) -> Ok (ix+1,a)
        with
        | _ -> Err "parseTable" 

        




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
//    TablesExtractor <| fun doc ix  -> 
//        if anchor.Position > ix then   
//           apply1 ma doc anchor.Position
//        else
//           Err "withSearchAnchor"

//let withSearchAnchorM (anchorQuery:TablesExtractor<SearchAnchor>) 
//                        (ma:TablesExtractor<'a>) : TablesExtractor<'a> = 
//    anchorQuery >>>= fun a -> withSearchAnchor a ma

//let advanceM (anchorQuery:TablesExtractor<SearchAnchor>) : TablesExtractor<unit> = 
//    withSearchAnchorM anchorQuery (dreturn ())


//let findText (search:string) (matchCase:bool) : TablesExtractor<Region> =
//    TablesExtractor <| fun doc ix  -> 
//        let range =  getRangeToEnd ix doc 
//        match boundedFind1 search matchCase extractRegion range with
//        | Some region -> Ok (ix, region)
//        | None -> Err <| sprintf "findText - '%s' not found" search

//let findTextStart (search:string) (matchCase:bool) : TablesExtractor<SearchAnchor> =
//    findText search matchCase |>>> startOfRegion

//let findTextEnd (search:string) (matchCase:bool) : TablesExtractor<SearchAnchor> =
//    findText search matchCase |>>> endOfRegion

///// Case sensitivity always appears to be true for Wildcard matches.
//let findPattern (search:string) : TablesExtractor<Region> =
//    TablesExtractor <| fun doc ix  -> 
//        let range =  getRangeToEnd ix doc 
//        match boundedFindPattern1 search extractRegion range with
//        | Some region -> Ok (ix, region)
//        | None -> Err <| sprintf "findPattern - '%s' not found" search
        
//let findPatternStart (search:string)  : TablesExtractor<SearchAnchor> =
//    findPattern search |>>> startOfRegion

//let findPatternEnd (search:string) : TablesExtractor<SearchAnchor> =
//    findPattern search |>>> endOfRegion

//let findTextMany (search:string) (matchCase:bool) : TablesExtractor<Region list> =
//    TablesExtractor <| fun doc ix  -> 
//        let range =  getRangeToEnd ix doc 
//        let finds = boundedFindMany search matchCase extractRegion range
//        Ok (ix,finds)

///// Case sensitivity always appears to be true for Wildcard matches.
//let findPatternMany (search:string) : TablesExtractor<Region list> =
//    TablesExtractor <| fun doc ix -> 
//        let range =  getRangeToEnd ix doc 
//        let finds = boundedFindPatternMany search extractRegion range
//        Ok (ix,finds)
