
module DocSoup.DocMonad

open System.Text.RegularExpressions
open Microsoft.Office.Interop

open DocSoup.Base


[<Struct>]
type State = State of string

type Ans<'a> = 
    | Err of string
    | Ok of State * 'a

let private ansMapM (fn:'a -> State -> Ans<'b>) (st0:State) (xs:'a list) : Ans<'b list> = 
    let rec work ac st ys = 
        match ys with
        | [] -> Ok(st, List.rev ac)
        | z :: zs -> 
            match fn z st with
            | Err msg -> Err msg
            | Ok(st1,a) -> work (a::ac) st1 zs
    work [] st0 xs


// DocMonad is Reader(immutable)+Reader+State+Error
// WARNING - region and state represent the same thing - must be kept in sync
type DocMonad<'a> = DocMonad of (Word.Document -> Region -> State -> Ans<'a>)


let inline apply1 (ma : DocMonad<'a>) (doc:Word.Document) (focus:Region) (st:State) : Ans<'a>= 
    let (DocMonad f) = ma in f doc focus st

let inline preturn (x:'a) : DocMonad<'a> = DocMonad <| fun _ _ st -> Ok(st,x)


let inline bindM (ma:DocMonad<'a>) (f : 'a -> DocMonad<'b>) : DocMonad<'b> =
    DocMonad <| fun doc focus st0 -> 
        match apply1 ma doc focus st0 with
        | Err msg -> Err msg
        | Ok(st1,a) -> apply1 (f a) doc focus st1

let forExprM (source: seq<'a>) (fn: 'a -> DocMonad<unit>) : DocMonad<unit> = 
    DocMonad <| fun doc focus st0 -> 
        let rec work st ys = 
            match ys with
            | [] -> Ok (st, ())
            | z :: zs -> 
                match apply1 (fn z) doc focus st with
                | Err msg -> Err msg
                | Ok (st1,a) -> work st1 zs
        work st0 (Seq.toList source)

let combineM (ma:DocMonad<unit>) (mb:DocMonad<unit>) : DocMonad<unit> = 
    DocMonad <| fun doc focus st0 -> 
        match apply1 ma doc focus st0 with
        | Err msg -> Err msg
        | Ok(st1,a) -> 
            match apply1 mb doc focus st1 with
            | Err msg -> Err msg
            | Ok(st2,a) -> Ok (st2, ())

let delayM (fn:unit -> DocMonad<'a>) : DocMonad<'a> = 
    bindM (preturn ()) fn 


let inline pzero () : DocMonad<'a> = 
    DocMonad <| fun _ _ _ -> Err "fail-pzero"

type DocMonadBuilder() = 
    member self.Return x        = preturn x
    member self.Bind (p,f)      = bindM p f
    member self.Zero ()         = pzero ()
    member self.For xs ma       = forExprM xs ma
    member self.Combine ma mb   = combineM ma mb
    member self.Delay fn        = delayM fn

let (docMonad:DocMonadBuilder) = new DocMonadBuilder()

let (>>=) (ma:DocMonad<'a>) (fn:'a -> DocMonad<'b>) : DocMonad<'b> = bindM ma fn


// Common monadic operations
let fmapM (fn:'a -> 'b) (ma:DocMonad<'a>) : DocMonad<'b> = 
    DocMonad <| fun doc focus st0 -> 
       match apply1 ma doc focus st0 with
       | Err msg -> Err msg
       | Ok(st1,a)-> Ok (st1, fn a)


let (|>>) (ma:DocMonad<'a>) (fn:'a -> 'b) : DocMonad<'b> = fmapM fn ma
let (<<|) (fn:'a -> 'b) (ma:DocMonad<'a>) : DocMonad<'b> = fmapM fn ma


let liftM (fn:'a -> 'x) (ma:DocMonad<'a>) : DocMonad<'x> = fmapM fn ma

let liftM2 (fn:'a -> 'b -> 'x) (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'x> = 
    docMonad { 
        let! a = ma
        let! b = mb
        return (fn a b)
    }

let liftM3 (fn:'a -> 'b -> 'c -> 'x) (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) : DocMonad<'x> = 
    docMonad { 
        let! a = ma
        let! b = mb
        let! c = mc
        return (fn a b c)
    }

let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) (md:DocMonad<'d>) : DocMonad<'x> = 
    docMonad { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        return (fn a b c d)
    }


let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) (md:DocMonad<'d>) (me:DocMonad<'e>) : DocMonad<'x> = 
    docMonad { 
        let! a = ma
        let! b = mb
        let! c = mc
        let! d = md
        let! e = me
        return (fn a b c d e)
    }

let tupleM2 (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'a * 'b> = 
    liftM2 (fun a b -> (a,b)) ma mb

let tupleM3 (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) : DocMonad<'a * 'b * 'c> = 
    liftM3 (fun a b c -> (a,b,c)) ma mb mc

let tupleM4 (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) (md:DocMonad<'d>) : DocMonad<'a * 'b * 'c * 'd> = 
    liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

let tupleM5 (ma:DocMonad<'a>) (mb:DocMonad<'b>) (mc:DocMonad<'c>) (md:DocMonad<'d>) (me:DocMonad<'e>) : DocMonad<'a * 'b * 'c * 'd * 'e> = 
    liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

let sequenceM (source:DocMonad<'a> list) : DocMonad<'a list> = 
    DocMonad <| fun doc focus st0 -> 
        let rec work ac st ys = 
            match ys with
            | [] -> Ok (st, List.rev ac)
            | ma :: zs -> 
                match apply1 ma doc focus st with
                | Err msg -> Err msg
                | Ok (st1,a) -> work (a::ac) st1 zs
        work [] st0 source

// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
let alt (ma:DocMonad<'a>) (mb:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus st0 ->
        match apply1 ma doc focus st0 with
        | Err _ -> apply1 mb doc focus st0
        | Ok (st1,a) -> Ok (st1, a)

let (<|>) (ma:DocMonad<'a>) (mb:DocMonad<'a>) : DocMonad<'a> = alt ma mb

// Applicative's (<*>)
let apM (mf:DocMonad<'a ->'b>) (ma:DocMonad<'a>) : DocMonad<'b> = 
    docMonad { 
        let! fn = mf
        let! a = ma
        return (fn a) 
    }

// Perform two actions in sequence. Ignore the results of the second action if both succeed.
let seqL (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'a> = 
    docMonad { 
        let! a = ma
        let! b = mb
        return a
    }

// Perform two actions in sequence. Ignore the results of the first action if both succeed.
let seqR (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'b> = 
    docMonad { 
        let! a = ma
        let! b = mb
        return b
    }

let (.>>) (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'a> = seqL ma mb
let (>>.) (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'b> = seqR ma mb

let count (ntimes:int) (ma:DocMonad<'a>) : DocMonad<'a []> = 
    DocMonad <| fun doc focus st0 ->
        let rec work ac ix st = 
            if ix < ntimes then                
                match apply1 ma doc focus st with
                | Err msg -> Err msg
                | Ok (st1,a) -> work (a :: ac) (ix+1) st1
            else 
                Ok (st, List.toArray <| List.rev ac)
        work [] 0 st0


let many (ma:DocMonad<'a>) : DocMonad<'a list> = 
    DocMonad <| fun doc focus st0 ->
        let rec work ac st = 
            match apply1 ma doc focus st with
            | Err _ -> Ok (st,List.rev ac)
            | Ok (st1,a) -> work (a :: ac) st1
        work [] st0

let many1 (ma:DocMonad<'a>) : DocMonad<'a list> = 
    docMonad { 
        let! a = ma
        let! rest = many ma
        return (a::rest)
    }


// DocMonad specific operations
let runOnFile (ma:DocMonad<'a>) (fileName:string) : Ans<'a> =
    if System.IO.File.Exists (fileName) then
        let app = new Word.ApplicationClass (Visible = true) 
        let doc = app.Documents.Open(FileName = ref (fileName :> obj))
        let region1 = maxRegion doc
        let text1 = regionText region1 doc
        let ans = apply1 ma doc region1 (State text1)
        doc.Close(SaveChanges = ref (box false))
        app.Quit()
        ans
    else Err <| sprintf "Cannot find file %s" fileName

let runOnFileE (ma:DocMonad<'a>) (fileName:string) : 'a =
    match runOnFile ma fileName with
    | Err msg -> failwith msg
    | Ok (_,a) -> a


let throwError (msg:string) : DocMonad<'a> = 
    DocMonad <| fun _ _  _ -> Err msg

let swapError (msg:string) (ma:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus st0 ->
        match apply1 ma doc focus st0 with
        | Err _ -> Err msg
        | Ok (st1,a) -> Ok (st1,a)

let (<?>) (ma:DocMonad<'a>) (msg:string) : DocMonad<'a> = swapError msg ma

let augmentError (fn:string -> string) (ma:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus st0 ->
        match apply1 ma doc focus st0 with
        | Err msg -> Err <| fn msg
        | Ok (st1,a) -> Ok (st1,a)

let withInput (fn:string -> Ans<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus st0 ->
        try
            match st0 with | State(s) -> fn s
        with    
        | ex -> Err "withInput"




// Get the text in the currently focused region.

// Returns the raw text that may include control characters.
// Shouldn't expose this...
let rawText : DocMonad<string> = 
    DocMonad <| fun doc focus st0 -> 
        match st0 with
        | State(s) -> Ok (st0,s)
            


// Removes all control characters except CR & LF.
let cleanText : DocMonad<string> = 
    fmapM (fun (s:string) -> Regex.Replace(s, @"[\p{C}-[\r\n]]+", "")) rawText



// Get the currently focused region.
let askFocus : DocMonad<Region> = 
    DocMonad <| fun _ focus st0 ->  
        Ok (st0,focus)

let asksFocus (fn:Region -> 'a) : DocMonad<'a> = 
    DocMonad <| fun doc focus st0 ->  
        Ok (st0, fn focus)

let local (project:Region -> Region) (ma:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus st0 ->  
        apply1 ma doc (project focus) st0





// Probably should not be part of the API...
let liftGlobalOperation (fn : Word.Document -> 'a) : DocMonad<'a> = 
    DocMonad <| fun doc _ st0 ->
        try
            Ok (st0, fn doc)
        with
        | ex -> Err <| ex.ToString()


let liftOperation (fn : Word.Range -> 'a) : DocMonad<'a> = 
    DocMonad <| fun doc focus st0 ->
        try
            let range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            Ok (st0, fn range)
        with
        | ex -> Err <| ex.ToString()


let optional (ma:DocMonad<'a>) : DocMonad<'a option> = 
    DocMonad <| fun doc focus st0 ->
        match apply1 ma doc focus st0 with
        | Err _ -> Ok (st0, None)
        | Ok (st1,a) -> Ok (st1, Some a)

let optionalz (ma:DocMonad<'a>) : DocMonad<unit> = 
    DocMonad <| fun doc focus st0 ->
        match apply1 ma doc focus st0 with
        | Err _ -> Ok (st0, ())
        | Ok (st1,_) -> Ok (st1, ())

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
    DocMonad <| fun doc focus st0 -> 
        try 
            let range0:Word.Range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            let table1:Word.Table = range0.Tables.[index]
            let range1:Word.Range = table1.Range
            apply1 ma doc (extractRegion range1) st0
        with
        | ex -> Err <| ex.ToString() 


// Needs a better name...
let mapTablesWith (ma:DocMonad<'a>) : DocMonad<'a list> = 
    DocMonad <| fun doc focus st0 -> 
        try 
            let range0:Word.Range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            let tables:Word.Table list = (range0.Tables |> Seq.cast<Word.Table> |> Seq.toList)
            ansMapM (fun table st -> let region = extractRegion (table :> Word.Table).Range in apply1 ma doc region st) st0 tables
        with
        | ex -> Err <| ex.ToString() 


// Strangely this appears to count from zero
let cell (row:int, col:int) (ma:DocMonad<'a>) : DocMonad<'a> = 
    DocMonad <| fun doc focus st0 -> 
        try 
            let range0:Word.Range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            let table1:Word.Table = range0.Tables.[1]
            let range1:Word.Range = table1.Cell(row,col).Range
            apply1 ma doc (extractRegion range1) st0
        with
        | ex -> Err <| ex.ToString() 


let mapCellsWith (ma:DocMonad<'a>) : DocMonad<'a list> = 
    DocMonad <| fun doc focus st0 -> 
        try 
            let range0:Word.Range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
            let cells:Word.Cell list = (range0.Cells |> Seq.cast<Word.Cell> |> Seq.toList)
            ansMapM (fun cell st -> let region = extractRegion (cell :> Word.Cell).Range in apply1 ma doc region st) st0 cells
        with
        | ex -> Err <| ex.ToString() 
        
