// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Rows = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal.ExtractMonad
    open DocSoup.Internal.Consume
    open DocSoup

    type RowsExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableRow []> 

    let (extractor:RowsExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableRow []>()

    type Extractor<'a> = ExtractMonad<'a, Wordprocessing.TableRow []> 

   
    let internal rowConsume = makeConsumeModule ()

    let getItem : Extractor<Wordprocessing.TableRow> = rowConsume.GetItem

    let getItems : int -> Extractor<Wordprocessing.TableRow []> = rowConsume.GetItems

    let position : Extractor<int> = rowConsume.Position

    let getInput : Extractor<Wordprocessing.TableRow []> = rowConsume.GetInput

    let inputCount : Extractor<int> = rowConsume.InputCount