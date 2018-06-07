// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocSoup.Monad2

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
        let app = new Word.ApplicationClass (Visible = true) :> Word.Application
        let doc = app.Documents.Open(FileName = ref (fileName :> obj))
        let region1 = maxRegion doc
        let ans = apply1 ma doc region1
        doc.Close(SaveChanges = rbox false)
        app.Quit()
        ans
    else Err <| sprintf "Cannot find file %s" fileName


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
        
let getFocusText : DocSoup<string> =
    DocSoup <| fun doc focus -> 
        let text = regionText focus doc 
        Ok text
        

let findText (search:string) : DocSoup<Region> =
    DocSoup <| fun doc focus  -> 
        let range1 = getRange focus doc
        range1.Find.ClearFormatting ()
        if range1.Find.Execute (FindText = rbox search) then
            Ok <| extractRegion range1
        else
            Err <| sprintf "findText - '%s' not found" search
