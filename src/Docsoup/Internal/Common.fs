﻿// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocSoup.Internal

module Common = 

    open System.IO
    open System.Text.RegularExpressions

    // Add references via the COM tab for Office and Word
    // All the PIA stuff online is outdated for Office 365 / .Net 4.5 / VS2015 
    open Microsoft.Office.Interop


    
    
    let rbox (v : 'a) : obj ref = ref (box v)

    let primitiveExtract (extract:Word.Document -> 'a) (fileName:string) : Result<'a,string> =
        if System.IO.File.Exists (fileName) then
            let app = new Word.ApplicationClass (Visible = false) :> Word.Application
            try 
                let doc = app.Documents.Open(FileName = rbox fileName)
                let ans = extract doc
                doc.Close(SaveChanges = rbox false)
                app.Quit()
                Ok ans
            with
              | ex -> 
                let msg = sprintf "exn type: %A\nmessage: %s" (ex.GetType())  ex.Message
                try 
                    app.Quit ()
                    Error msg
                with
                | _ -> Error msg
        else 
            Error <| sprintf "Cannot find file: %s" fileName



    let rangeText (range:Word.Range) : string = range.Text

    let cleanRangeText (range:Word.Range) : string = 
        let str = Regex.Replace(range.Text, @"[\p{C}-[\r\n]]+", "")
        str.Trim() 


    let find1 (search:string) (matchCase:bool) (range:Word.Range) : Word.Range option = 
        range.Find.ClearFormatting ()
        let found = 
            range.Find.Execute (FindText = rbox search, 
                                MatchWildcards = rbox false,
                                MatchCase = rbox matchCase,
                                Forward = rbox true) 
        if found then Some range else None


    let findMany (search:string) (matchCase:bool) (range:Word.Range) : Word.Range list = 
        let rec work (current:Word.Range) (cont: Word.Range list -> Word.Range list) = 
        
            printfn "Initial Range { start: %i; end %i} " range.Start range.End
            let found = 
                current.Find.Execute (FindText = rbox search, 
                                    MatchWildcards = rbox false,
                                    MatchCase = rbox matchCase,
                                    Forward = rbox true) 
            if found then 
                let ans = current
                work current ( fun xs -> cont (current :: xs))
            else 
                cont []


        range.Find.ClearFormatting ()
        work range (fun xs -> xs)


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

    let getRange (region:Region) (doc:Word.Document) : Word.Range = 
        doc.Range(rbox <| region.RegionStart, rbox <| region.RegionEnd - 1)

    let getRangeToEnd (pos:int) (doc:Word.Document) : Word.Range = 
        doc.Range(rbox <| pos, rbox <| doc.Range().End)

    // Use (Single Case) Struct Unions to get the same things as Haskell Newtypes.

    let isSubregionOf (haystack:Region) (needle:Region) : bool = 
        needle.RegionStart >= haystack.RegionStart && needle.RegionEnd <= haystack.RegionEnd

    let regionText (focus:Region) (doc:Word.Document) : string = 
        let fStart = focus.RegionStart
        let fEnd = focus.RegionEnd
        cleanRangeText <| doc.Range(rbox fStart, rbox fEnd)
    

    let regionPlus (r1:Region) (r2:Region) : Region = 
        { RegionStart = min r1.RegionStart r2.RegionStart
          RegionEnd = max r1.RegionEnd r2.RegionEnd }

    let regionConcat (regions:Region list) : Region option = 
        match regions with
        | [] -> None
        | (r::rs) -> Some <| List.fold regionPlus r rs


    // cells


    let tryFindCell (predicate:Word.Cell -> bool) (table:Word.Table) : option<Word.Cell> = 
        let rowMax = table.Rows.Count 
        let colMax = table.Columns.Count
        let softCell (table:Word.Table) (row:int) (col:int) : option<Word.Cell> = 
            try 
                Some <| table.Cell(row, col)
            with
            | _ -> None

        let rec work (row:int) (col:int) : option<Word.Cell> = 
            if row >= rowMax then 
                None
            else
                if col >= colMax then 
                    work (row+1) 0 
                else
                    match softCell table row col with
                    | None -> work row (col+1)
                    | Some cell ->
                        if predicate cell then
                            Some cell
                        else work row (col+1)
        work 0 0                 



    /// Design note - the Word automation API provides a table ``ID`` but it 
    /// cannot be used on for lookup.
    type TableAnchor = 
        internal 
            { TableIndex: int }

        /// Because tables are 1-indexed in word the default constructor is called 
        /// ``First`` rather than ``Zero``.
        static member internal First = { TableIndex = 1 }
        member internal x.Index = x.TableIndex
        member internal x.Next = { TableIndex = x.TableIndex + 1 }
        member internal x.Prev = { TableIndex = x.TableIndex - 1 }

    let internal getTable (anchor:TableAnchor) (doc:Word.Document) : option<Word.Table> = 
        try
            Some <| doc.Range().Tables.Item(anchor.Index)
        with
        | _ -> None 


    /// Fully user visible
    type CellIndex = 
        { RowIx: int
          ColIx : int }

        /// Because tables are 1-indexed in word the default constructor is called 
        /// ``First`` rather than ``Zero``.
        static member internal First = { RowIx = 1; ColIx = 1 }
        member v.IncrRow = { v with RowIx = v.RowIx + 1 }
        member v.DecrRow = { v with RowIx = v.RowIx - 1 }
        member v.IncrCol = { v with ColIx = v.ColIx + 1 }
        member v.DecrCol = { v with ColIx = v.ColIx - 1 }

    type CellAnchor = 
        internal 
            { TableIx: TableAnchor
              CellIx: CellIndex }
        member internal x.TableAnchor : TableAnchor = x.TableIx
        member internal x.Row : int = x.CellIx.RowIx
        member internal x.Column : int = x.CellIx.ColIx


    let internal getCell (anchor:CellAnchor) (doc:Word.Document) : option<Word.Cell> = 
        try
            let table : Word.Table = doc.Range().Tables.Item(anchor.TableAnchor.Index)
            let cell : Word.Cell = table.Cell(anchor.Row, anchor.Column)
            Some cell
        with
        | _ -> None 

    let internal firstCell (table:TableAnchor)  = 
        { TableIx = table; CellIx = CellIndex.First }
    



    /// Word: Range.Find
    /// Operationally this is quite confusing.
    /// The first success must be withing the range search, e.g:
    /// > doc.Tables(6).Range.Find.Execute (FindText = ...
    /// but successive iterations apparently look after the range found, this
    /// means that results are not bound to be within the initial range.

    let boundedFind1 (search:string) (matchCase:bool) (mapper:Word.Range -> 'a) (initialRange:Word.Range) : 'a option = 
        let regionMax = extractRegion initialRange
        let range = initialRange
        range.Find.ClearFormatting ()
        let found = 
            range.Find.Execute (FindText = rbox search, 
                                MatchWildcards = rbox false,
                                MatchCase = rbox matchCase,
                                Forward = rbox true) 
        if found && isSubregionOf regionMax (extractRegion range) then
            Some <| mapper range
        else None


    let boundedFindPattern1 (search:string)  (mapper:Word.Range -> 'a) (initialRange:Word.Range) : 'a option = 
        let regionMax = extractRegion initialRange
        let range = initialRange
        range.Find.ClearFormatting ()
        let found = 
            range.Find.Execute (FindText = rbox search, 
                                MatchWildcards = rbox true,
                                MatchCase = rbox true,
                                Forward = rbox true) 
        if found && isSubregionOf regionMax (extractRegion range) then
            Some <| mapper range
        else None

    let boundedFindMany (search:string) (matchCase:bool) (mapper:Word.Range -> 'a) (initialRange:Word.Range) : 'a list = 
        let regionMax = extractRegion initialRange
        let producer (range:Word.Range) : ('a * Word.Range) option = 
            let found = 
                range.Find.Execute (FindText = rbox search, 
                                    MatchWildcards = rbox false,
                                    MatchCase = rbox matchCase,
                                    Forward = rbox true) 
            if found && isSubregionOf regionMax (extractRegion range) then
                Some (mapper range, range)    
            else None
        initialRange.Find.ClearFormatting ()
        List.unfold producer initialRange


    let boundedFindPatternMany (search:string) (mapper:Word.Range -> 'a) (initialRange:Word.Range) : 'a list = 
        let regionMax = extractRegion initialRange
        let producer (range:Word.Range) : ('a * Word.Range) option = 
            let found = 
                range.Find.Execute (FindText = rbox search, 
                                    MatchWildcards = rbox true,
                                    MatchCase = rbox true,
                                    Forward = rbox true) 
            if found && isSubregionOf regionMax (extractRegion range) then
                Some (mapper range, range)    
            else None
        initialRange.Find.ClearFormatting ()
        List.unfold producer initialRange


