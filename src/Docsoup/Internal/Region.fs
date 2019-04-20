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
    let isProperRegion (r1:Region) : bool = r1.Start >= 0 && r1.Start <= r1.End

    let intersection (r1:Region) (r2:Region) : Region option = 
        let (first,second) = if r1.Start <= r2.Start then (r1,r2) else (r2, r1)
        if second.Start <= first.End then 
            Some { Start = second.Start; End = min first.End second.End  }
         else None

        


    /// Find the (improper) union of two regions.
    /// If the regions do not overlap, the union will include the whole in the middle.
    let union (r1:Region) (r2:Region) : Region = 
        { Start = min r1.Start r2.Start
          End =   max r1.End   r2.End }

    let concat (regions:Region list) : Region option = 
        match regions with
        | [] -> None
        | (r::rs) -> Some <| List.fold union r rs

