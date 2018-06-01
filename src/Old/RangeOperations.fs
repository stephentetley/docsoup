module Old.RangeOperations


// Add references via the COM tab for Office and Word
// All the PIA stuff online is outdated for Office 365 / .Net 4.5 / VS2015 
open Microsoft.Office.Interop

 

let rightBorder (rbase : Word.Range) (anchor : Word.Range) : int = 
    rbase.End - anchor.End

let leftBorder (rbase : Word.Range) (anchor : Word.Range) : int = 
    anchor.Start - rbase.Start

/// Predicates

let equals (ra : Word.Range) (rb : Word.Range) : bool = 
    ra.Start = rb.Start && ra.End = rb.End

let contains (ra : Word.Range) (rb : Word.Range) : bool = 
    ra.Start <= rb.Start && ra.End >= rb.End

// Opposite of contains (if a contains b, b isIn a)
let isIn (ra : Word.Range) (rb : Word.Range) : bool = 
    ra.Start >= rb.Start && ra.End <= rb.End


let before (ra : Word.Range) (rb : Word.Range) : bool = 
    ra.End < rb.Start

let after (ra : Word.Range) (rb :Word.Range) : bool = 
    ra.Start > rb.End

// |---------|......... (a)
// ....|----------|.... (b)
//
// (a) overlaps-start of (b)
let overlapsStart (ra : Word.Range) (rb :Word.Range) : bool = 
    ra.End >= rb.Start && ra.End <= rb.End


// |---------|......... (a)
// ....|----------|.... (b)
//
// (b) overlaps-end of (a)
let overlapsEnd (ra : Word.Range) (rb :Word.Range) : bool = 
    ra.Start >= rb.Start && ra.Start <= rb.End


/// Set operations

// "Right Difference" - todo - write this without calling other functions
// |=========|......... (a)
// ....|==========|.... (b)
// ..........|====|.... (a) `rightDifference` (b)
//
//
// ....|==========|.... (a)
// |=========|......... (b)
// ..........|====|.... (a) `rightDifference` (b) ??
//
// Note - should we care about arg order?

let rightDifference (ra : Word.Range) (rb : Word.Range) : Option<Word.Range> = 
    if ra.End >= rb.Start && ra.End < rb.End  then 
        let newrange = rb.Duplicate
        newrange.Start <- ra.End + 1        // plus 1? ...
        Some <| newrange
    else None

// If we are confused about whether arg order is important it implies we have the wrong API...

// todo - write this without calling other functions
let leftDifference (rbase : Word.Range) (anchor : Word.Range) : Option<Word.Range> = 
    if leftBorder rbase anchor > 0 && rightBorder rbase anchor >= 0 then 
        let newrange = rbase.Duplicate
        newrange.End <- anchor.Start - 1        // minus 1? ...
        Some <| newrange
    else None

