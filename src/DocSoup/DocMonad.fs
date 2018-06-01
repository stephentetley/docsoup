
module DocSoup.DocMonad

open System.Text.RegularExpressions
open Microsoft.Office.Interop

open DocSoup.Base




// DocMonad is Reader(immutable)+Reader+Error
type DocMonad<'a> = DocMonad of (Word.Document -> Region -> Ans<'a>)


let inline apply1 (ma : DocMonad<'a>) (doc:Word.Document) (focus:Region) : Ans<'a>= 
    let (DocMonad f) = ma in f doc focus

let private unitM (x:'a) : DocMonad<'a> = DocMonad <| fun _ _ -> Ok x


let bindM (ma:DocMonad<'a>) (f : 'a -> DocMonad<'b>) : DocMonad<'b> =
    DocMonad <| fun doc focus -> 
        match apply1 ma doc focus with
        | Err msg -> Err msg
        | Ok a -> apply1 (f a) doc focus



type DocMonadBuilder() = 
    member self.Return x    = unitM x
    member self.Bind (p,f) = bindM p f
    member self.Zero ()     = unitM ()

let (docMonad:DocMonadBuilder) = new DocMonadBuilder()

// Common monadic operations
let fmapM (fn:'a -> 'b) (ma:DocMonad<'a>) : DocMonad<'b> = 
    DocMonad <| fun doc focus -> 
        Base.fmapM fn (apply1 ma doc focus)


let liftM (fn:'a -> 'x) (ma:DocMonad<'a>) : DocMonad<'x> = fmapM fn ma

let liftM2 (fn:'a -> 'b -> 'x) (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'x> = 
    DocMonad <| fun doc focus -> 
        match apply1 ma doc focus with
        | Err msg -> Err msg
        | Ok a -> 
            match apply1 mb doc focus with 
            | Err msg -> Err msg
            | Ok b -> Ok <| fn a b

let liftM3 (fn:'a -> 'b -> 'c -> 'x) (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) : DocMonad<'x> = 
    DocMonad <| fun doc focus -> 
        match apply1 ma doc focus with
        | Err msg -> Err msg
        | Ok a -> 
            match apply1 mb doc focus with 
            | Err msg -> Err msg
            | Ok b -> 
                match apply1 mc doc focus with 
                | Err msg -> Err msg
                | Ok c -> Ok <| fn a b c

let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) (md:DocMonad<'d>) : DocMonad<'x> = 
    DocMonad <| fun doc focus -> 
        match apply1 ma doc focus with
        | Err msg -> Err msg
        | Ok a -> 
            match apply1 mb doc focus with 
            | Err msg -> Err msg
            | Ok b -> 
                match apply1 mc doc focus with 
                | Err msg -> Err msg
                | Ok c -> 
                    match apply1 md doc focus with 
                    | Err msg -> Err msg
                    | Ok d -> Ok <| fn a b c d


let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) (md:DocMonad<'d>) (me:DocMonad<'e>) : DocMonad<'x> = 
    DocMonad <| fun doc focus -> 
        match apply1 ma doc focus with
        | Err msg -> Err msg
        | Ok a -> 
            match apply1 mb doc focus with 
            | Err msg -> Err msg
            | Ok b -> 
                match apply1 mc doc focus with 
                | Err msg -> Err msg
                | Ok c -> 
                    match apply1 md doc focus with 
                    | Err msg -> Err msg
                    | Ok d -> 
                        match apply1 me doc focus with 
                        | Err msg -> Err msg
                        | Ok e -> Ok <| fn a b c d e

let tupleM2 (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'a * 'b> = 
    liftM2 (fun a b -> (a,b)) ma mb

let tupleM3 (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) : DocMonad<'a * 'b * 'c> = 
    liftM3 (fun a b c -> (a,b,c)) ma mb mc

let tupleM4 (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) (md:DocMonad<'d>) : DocMonad<'a * 'b * 'c * 'd> = 
    liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

let tupleM5 (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) (md:DocMonad<'d>) (me:DocMonad<'e>) : DocMonad<'a * 'b * 'c * 'd * 'e> = 
    liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

let sequenceM (source:DocMonad<'a> list) : DocMonad<'a list> = 
    DocMonad <| fun doc focus -> 
        Base.sequenceM <| List.map (fun fn -> apply1 fn doc focus) source

// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
let alt (ma:DocMonad<'a>) (mb:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus ->
        match apply1 ma doc focus with
        | Err _ -> apply1 mb doc focus
        | Ok a -> Ok a

// Applicative's (<*>)
let apM (mf:DocMonad<'a ->'b>) (ma:DocMonad<'a>) : DocMonad<'b> = 
    DocMonad <| fun doc focus ->
        match apply1 mf doc focus with
        | Err msg -> Err msg
        | Ok fn -> 
            match apply1 ma doc focus with
            | Err msg -> Err msg
            | Ok a -> Ok <| fn a

// Perform two actions in sequence. Ignore the results of the second action if both succeed.
let seqL (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'a> = 
    DocMonad <| fun doc focus ->
        match apply1 ma doc focus with
        | Err msg -> Err msg
        | Ok a -> 
            match apply1 mb doc focus with
            | Err msg -> Err msg
            | Ok _ -> Ok a

// Perform two actions in sequence. Ignore the results of the first action if both succeed.
let seqR (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'b> = 
    DocMonad <| fun doc focus ->
        match apply1 ma doc focus with
        | Err msg -> Err msg
        | Ok _ -> 
            match apply1 mb doc focus with
            | Err msg -> Err msg
            | Ok b -> Ok b

// DocMonad specific operations
let runOnFile (ma:DocMonad<'a>) (fileName:string) : Ans<'a> =
    if System.IO.File.Exists (fileName) then
        let app = new Word.ApplicationClass (Visible = true) 
        let doc = app.Documents.Open(FileName = ref (fileName :> obj))
        let ans = apply1 ma doc (maxRegion doc)
        doc.Close(SaveChanges = ref (box false))
        app.Quit()
        ans
    else Err <| sprintf "Cannot find file %s" fileName

let runOnFileE (ma:DocMonad<'a>) (fileName:string) : 'a =
    match runOnFile ma fileName with
    | Err msg -> failwith msg
    | Ok a -> a


let throwError (msg:string) : DocMonad<'a> = 
    DocMonad <| fun _ _  -> Err msg

let swapError (msg:string) (ma:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus ->
        match apply1 ma doc focus with
        | Err msg -> Err msg
        | Ok result -> Ok result


let augmentError (fn:string -> string) (ma:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus ->
        match apply1 ma doc focus with
        | Err msg -> Err <| fn msg
        | Ok result -> Ok result


// Get the text in the currently focused region.

// Returns the raw text that may include control characters.
let rawText : DocMonad<string> = 
    DocMonad <| fun doc focus -> 
        try
            let range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            Ok <| range.Text
        with
        | ex -> Err <| sprintf "text: %s" (ex.ToString())


// Removes all control characters except CR & LF.
let cleanText : DocMonad<string> = 
    fmapM (fun (s:string) -> Regex.Replace(s, @"[\p{C}-[\r\n]]+", "")) rawText



// Get the currently focused region.
let askFocus : DocMonad<Region> = 
    DocMonad <| fun doc focus ->  
        Ok <| focus

let asksFocus (fn:Region -> 'a) : DocMonad<'a> = 
    DocMonad <| fun doc focus ->  
        Ok <| fn focus

let local (project:Region -> Region) (ma:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus ->  
        apply1 ma doc (project focus) 



// Probably should not be part of the API...
let liftGlobalOperation (fn : Word.Document -> 'a) : DocMonad<'a> = 
    DocMonad <| fun doc _ ->
        try
            Ok <| fn doc
        with
        | ex -> Err <| ex.ToString()


let liftOperation (fn : Word.Range -> 'a) : DocMonad<'a> = 
    DocMonad <| fun doc focus ->
        try
            let range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            Ok <| fn range
        with
        | ex -> Err <| ex.ToString()


let optional (ma:DocMonad<'a>) : DocMonad<'a option> = 
    DocMonad <| fun doc focus ->
        match apply1 ma doc focus with
        | Err _ -> Ok None
        | Ok a -> Ok <| Some a

let optionalz (ma:DocMonad<'a>) : DocMonad<unit> = 
    DocMonad <| fun doc focus ->
        match apply1 ma doc focus with
        | Err _ -> Ok ()
        | Ok _ -> Ok ()

// Range delimited.
let countTables : DocMonad<int> = 
    liftOperation <| fun doc -> doc.Tables.Count


// Range delimited.
let countSections : DocMonad<int> = 
    liftOperation <| fun rng -> rng.Sections.Count

// Range delimited.
let countCells : DocMonad<int> = 
    liftOperation <| fun rng -> rng.Cells.Count

// Range delimited.
let countParagraphs : DocMonad<int> = 
    liftOperation <| fun rng -> rng.Paragraphs.Count

// Range delimited.
let countCharacters : DocMonad<int> = 
    liftOperation <| fun rng -> rng.Characters.Count

let table (index:int) (ma:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus -> 
        try 
            let range0:Word.Range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            let table1:Word.Table = range0.Tables.[index]
            let range1:Word.Range = table1.Range
            apply1 ma doc (extractRegion range1)
        with
        | ex -> Err <| ex.ToString() 


// Needs a better name...
let mapTablesWith (ma:DocMonad<'a>) : DocMonad<'a list> = 
    DocMonad <| fun doc focus -> 
        try 
            let range0:Word.Range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            let tables:Word.Table list = (range0.Tables |> Seq.cast<Word.Table> |> Seq.toList)
            ansMapM (fun table -> let region = extractRegion (table :> Word.Table).Range in apply1 ma doc region) tables
        with
        | ex -> Err <| ex.ToString() 


// Strangely this appears to count from zero
let cell (row:int, col:int) (ma:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus -> 
        try 
            let range0:Word.Range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            let table1:Word.Table = range0.Tables.[1]
            let range1:Word.Range = table1.Cell(row,col).Range
            apply1 ma doc (extractRegion range1)
        with
        | ex -> Err <| ex.ToString() 


let mapCellsWith (ma:DocMonad<'a>) : DocMonad<'a list> = 
    DocMonad <| fun doc focus -> 
        try 
            let range0:Word.Range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            let cells:Word.Cell list = (range0.Cells |> Seq.cast<Word.Cell> |> Seq.toList)
            ansMapM (fun cell -> let region = extractRegion (cell :> Word.Cell).Range in apply1 ma doc region) cells
        with
        | ex -> Err <| ex.ToString() 
        
