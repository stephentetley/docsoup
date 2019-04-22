// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module TableExtractor = 
    
    open DocumentFormat.OpenXml

    open DocSoup

    type TableExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Table> 

    let (tableExtractor:TableExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Table>()

    type TableExtractor<'a> = ExtractMonad<Wordprocessing.Table,'a> 


    let rows : TableExtractor<seq<Wordprocessing.TableRow>> = 
        asks (fun table -> table.Elements<Wordprocessing.TableRow>())

    let row (index:int) : TableExtractor<Wordprocessing.TableRow> = 
        tableExtractor { 
            let! xs = rows
            return! liftOption (Seq.tryItem index xs)
        }

    let tableInnerText : TableExtractor<string> = 
        asks (fun table -> table.InnerText)