// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocSoup.Base

open System.IO
open System.Text.RegularExpressions

// Add references via the COM tab for Office and Word
// All the PIA stuff online is outdated for Office 365 / .Net 4.5 / VS2015 
open Microsoft.Office.Interop



let rbox (v : 'a) : obj ref = ref (box v)



// StringReader appears to be the best way of doing this. 
// Trying to split on a character (or character combo e.g. "\r\n") seems unreliable.
let sRestOfLine (s:string) : string = 
    use reader = new StringReader(s)
    reader.ReadLine ()


// Range is a very heavy object to be manipulating start and end points
// Use an alternative...
[<StructuredFormatDisplay("Region: {RegionStart} to {RegionEnd}")>]
type Region = { RegionStart : int; RegionEnd : int}


let extractRegion (range:Word.Range) : Region = { RegionStart = range.Start; RegionEnd = range.End }
    
let maxRegion (doc:Word.Document) : Region = extractRegion <| doc.Range()

let getRange (region:Region)  (doc:Word.Document) : Word.Range = 
    doc.Range(rbox <| region.RegionStart, rbox <| region.RegionEnd - 1)

let isSubregionOf (major:Region) (minor:Region) : bool = 
    minor.RegionStart >= major.RegionStart && minor.RegionEnd <= major.RegionEnd

let regionText (focus:Region) (doc:Word.Document) : string = 
    let range = doc.Range(rbox <| focus.RegionStart, rbox <| focus.RegionEnd)
    Regex.Replace(range.Text, @"[\p{C}-[\r\n]]+", "")

