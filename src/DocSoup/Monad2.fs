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
    member self.Return x            = preturn x
    member self.Bind (p,f)          = bindM p f
    member self.Zero ()             = szero ()
    // member self.For (xs,ma)         = forExprM xs ma
    member self.Combine (ma,mb)     = combineM ma mb
    member self.Delay fn            = delayM fn

 // Prefer "parse" to "parser" for the _Builder instance

let (docSoup:DocSoupBuilder) = new DocSoupBuilder()

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

let fparse (p:TextParser<'a>) : DocSoup<'a> = 
    DocSoup <| fun doc focus -> execFParsec doc focus p 

let findText (search:string) : DocSoup<Region> =
    DocSoup <| fun doc focus  -> 
        let range1 = getRange focus doc
        range1.Find.ClearFormatting ()
        if range1.Find.Execute (FindText = rbox search) then
            Ok <| extractRegion range1
        else
            Err <| sprintf "findText - '%s' not found" search
