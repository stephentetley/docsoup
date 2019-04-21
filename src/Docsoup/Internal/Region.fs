// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocSoup.Internal

[<RequireQualifiedAccess>]
module Region = 


    open Microsoft.Office.Interop

    open DocSoup.Internal.Common

    /// Range is a very heavy object for just be manipulating start and end points
    /// Use an alternative...
    type Region = { Start : int; End : int}

    let ofRange (range:Word.Range) : Region = { Start = range.Start; End = range.End }

    /// Extract a range of a document contained within the region.
    let extractRange (doc:Word.Document) (region:Region) : Word.Range = 
        doc.Range(rbox <| region.Start, rbox <| region.End - 1)


    /// Is the region well formed?
    let properRegion (r1:Region) : bool = r1.Start >= 0 && r1.Start <= r1.End

    /// Is r1 before r2 (and does not overlap)?
    let before (r1:Region) (r2:Region) : bool = 
        r1.End <= r2.Start

    /// Is r1 after r2 (and does not overlap)?
    let after (r1:Region) (r2:Region) : bool = 
        r1.Start >= r2.End

    /// Does r1 contain r2?
    let contains (r1:Region) (r2:Region) : bool = 
        r1.Start <= r2.Start && r1.End >= r2.End


    /// Word's find is horrible to use because uses mutation.
    let find1 (search:string) (matchCase:bool) (range:Word.Range) : Region option = 
        range.Find.ClearFormatting ()
        let found = 
            range.Find.Execute (FindText = rbox search, 
                                MatchWildcards = rbox false,
                                MatchCase = rbox matchCase,
                                Forward = rbox true) 
        if found then Some (ofRange range) else None

    /// Word's find is horrible to use because uses mutation.
    let findMany (search:string) (matchCase:bool) (range:Word.Range) : Region list = 
        let rec work (current:Word.Range) (cont: Region list -> Region list) = 
            let found = 
                current.Find.Execute (FindText = rbox search, 
                                    MatchWildcards = rbox false,
                                    MatchCase = rbox matchCase,
                                    Forward = rbox true) 
            if found then 
                printfn "Found {Start = %i ; End = %i}" current.Start current.End
                let r1 = ofRange current
                work current ( fun xs -> cont (r1 :: xs))
            else 
                cont []
        range.Find.ClearFormatting ()
        work range (fun xs -> xs)

        


    ///// Find the (improper) union of two regions.
    ///// If the regions do not overlap, the union will include the whole in the middle.
    //let union (r1:Region) (r2:Region) : Region = 
    //    { Start = min r1.Start r2.Start
    //      End =   max r1.End   r2.End }

    //let concat (regions:Region list) : Region option = 
    //    match regions with
    //    | [] -> None
    //    | (r::rs) -> Some <| List.fold union r rs

