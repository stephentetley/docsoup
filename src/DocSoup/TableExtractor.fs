// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup


module TableExtractor = 
    
    open DocumentFormat.OpenXml

    open DocSoup

    type TableExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Table> 

    let (tableExtractor:TableExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Table>()

    type TableExtractor<'a> = ExtractMonad<Wordprocessing.Table,'a> 

    let tableInnerText : TableExtractor<string> = 
        asks (fun table -> table.InnerText)