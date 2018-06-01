
module Old.Extractors

// Add references via the COM tab for Office and Word
// All the PIA stuff online is outdated for Office 365 / .Net 4.5 / VS2015 
open Microsoft.Office.Interop

open DocSoup.Base
open Old.RangeOperations

type Result<'a> = 
    | Okay of 'a
    | Fail of string

    
// Design - what we would like are delimited regions like the Reader monad's @local.
// We can achieve this but "almost all the work" goes into the first branch of @local
// (i.e the projection: WRange -> WRange) rather than the second branch (the monadic action).
// This means we lose the monadic utilies like (<|>) in the projection.
//
// The State monad is not a good alternative as updates are undelimited - an update changes 
// state for the rest of the computation.
//
// It looks like a two level EDSL is going to be the best approach - one level is monadic
// and handles binding, feeding the doc (Reader) and error handling. The other (inner) level 
// provides some sort of path expressions on Ranges that are run by the outer level.

type WRange = Word.Range
    
        

// Extractor is Reader+Error

type Extractor<'a> = Extractor of (WRange -> Result<'a>)



let fail : Extractor<'a> = Extractor (fun rng -> Fail "fail")

let apply1 (p : Extractor<'a>) (rng : WRange) : Result<'a> = 
    let (Extractor pf) = p in pf rng


let unit (x : 'a) : Extractor<'a> = 
    Extractor (fun rng -> Okay x)


let bind (p : Extractor<'a>) (f : 'a -> Extractor<'b>) : Extractor<'b> =
    Extractor <| fun rng -> 
        match apply1 p rng with
        | Okay a -> apply1 (f a) rng
        | Fail msg -> Fail msg


let fmap (f : 'a -> 'b) (p : Extractor<'a>) : Extractor<'b> = 
    bind p (unit << f)

type ExtractorBuilder() = 
    member self.Return x = unit x
    member self.Bind (p,f) = bind p f
    member self.Zero () = fail


let extractor = new ExtractorBuilder()

let empty : Extractor<'a> = fail


let delimit (upd : WRange -> WRange) (p : Extractor<'a>) : Extractor<'a> = 
    Extractor <| fun rng ->
        let rng1 = rng.Duplicate
        apply1 p (upd rng1)


// Left-biased
let alt (p : Extractor<'a>) (q : Extractor<'a>) : Extractor<'a> =
    Extractor <| fun rng -> 
        match apply1 p rng with
        | Okay a -> Okay a
        | Fail _ -> apply1 q rng



let (<|>) (p : Extractor<'a>) (q : Extractor<'a>) : Extractor<'a> = alt p q


let ap (p : Extractor<'a -> 'b>) (q : Extractor<'a>) : Extractor<'b> =
    extractor { let! f = p
                let! a = q
                return (f a)
                }

let (<*>) (p : Extractor<'a -> 'b>) (q : Extractor<'a>) : Extractor<'b> = ap p q

let apLeft (p : Extractor<'a>) (q : Extractor<'b>) : Extractor<'a> = 
    extractor { let! a = p
                let! _ = q
                return a
                }

let apRight (p : Extractor<'a>) (q : Extractor<'b>) : Extractor<'b> = 
    extractor { let! _ = p
                let! a = q
                return a
                }

let ( *> ) (p : Extractor<'a>) (q : Extractor<'b>) : Extractor<'b> = apRight p q
let ( <* ) (p : Extractor<'a>) (q : Extractor<'b>) : Extractor<'a> = apLeft p q


let text : Extractor<string> = 
    Extractor <| fun rng -> 
        match rng with
        | null -> Fail "Range is null"
        | rng1 -> Okay <| rng1.Text


// Note - the index is local within the range
// Also indexing is from 1 (must check this...)
let withTable (i:int) (p : Extractor<'a>) : Extractor<'a> =
    Extractor <| fun rng -> 
        if i < rng.Tables.Count then 
            let rng1 = rng.Tables.Item(i).Range
            apply1 p rng1
        else Fail "Table out of Range"

// look for line end...
let restOfLine : Extractor<string> = 
    Extractor <| fun rng -> 
        match rng with
        | null -> Fail "restOfLine - range is null"
        | rng1 -> Okay <| sRestOfLine rng1.Text


// To check - does duplicating range work as expected...
let find (s:string) : Extractor<string> = 
    let upd (rng : WRange) = 
        rng.Find.ClearFormatting ()
        let ans = rng.Find.Execute(FindText = rbox s)
        rng
    delimit upd text


//    let findr (s:string) (p:Extractor<'a>) : Extractor<'a> = 
//        let upd (rng : WRange) = 
//            rng.Find.ClearFormatting ()
//            let ans = rng.Find.Execute(FindText = rbox s)
//            rng
//        delimit upd p


let test (p : Extractor<'a>) (filepath : string) : 'a = 
    let app = new Word.ApplicationClass (Visible = true) 
    let doc = app.Documents.Open(FileName = rbox filepath)
    let dstart = doc.Content.Start
    let dend = doc.Content.End
    let rng = doc.Range(rbox dstart, rbox dend)
    let ans = apply1 p rng 
    doc.Close(SaveChanges = rbox false)
    app.Quit()
    match ans with
    | Fail msg -> failwith msg
    | Okay a -> a


