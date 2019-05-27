// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Tables = 
    
    open DocumentFormat.OpenXml

    
    open DocSoup.Internal
    open DocSoup.Internal.ExtractMonad
    open DocSoup.Internal.Consume

    type TablesExtractorBuilder = ExtractMonadBuilder<Wordprocessing.Table []> 

    let (extractor:TablesExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.Table []>()

    type Extractor<'a> = ExtractMonad<'a, Wordprocessing.Table []> 


    let internal tableConsume = makeConsumeModule ()

    let getItem : Extractor<Wordprocessing.Table> = tableConsume.GetItem

    let getItems : int -> Extractor<Wordprocessing.Table []> = tableConsume.GetItems

    let position : Extractor<int> = tableConsume.Position

    let getInput : Extractor<Wordprocessing.Table []> = tableConsume.GetInput

    let inputCount : Extractor<int> = tableConsume.InputCount

    let satisfy : (Wordprocessing.Table -> bool) -> Extractor<Wordprocessing.Table> = tableConsume.Satisfy