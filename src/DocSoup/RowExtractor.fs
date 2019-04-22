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

    let rowInnerText : RowExtractor<string> = 
        asks (fun row -> row.InnerText)