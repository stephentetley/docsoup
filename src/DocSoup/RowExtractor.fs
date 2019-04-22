// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module RowExtractor = 
    
    open DocumentFormat.OpenXml

    open DocSoup

    type RowExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableRow> 

    let (rowExtractor:RowExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableRow>()

    type RowExtractor<'a> = ExtractMonad<Wordprocessing.TableRow,'a> 

    
    let cells : RowExtractor<seq<Wordprocessing.TableCell>> = 
        asks (fun row -> row.Elements<Wordprocessing.TableCell>())

    let cell (index:int) : RowExtractor<Wordprocessing.TableCell> = 
        rowExtractor { 
            let! xs = cells
            return! liftOption (Seq.tryItem index xs)
        }


    let rowInnerText : RowExtractor<string> = 
        asks (fun row -> row.InnerText)