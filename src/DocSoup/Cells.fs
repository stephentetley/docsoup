// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<RequireQualifiedAccess>]
module Cells = 
    
    open System.Text
    open System.Text.RegularExpressions

    open DocumentFormat.OpenXml

    open DocSoup.Internal
    open DocSoup.Internal.ExtractMonad
    open DocSoup.Internal.Consume

    type CellsExtractorBuilder = ExtractMonadBuilder<Wordprocessing.TableCell []> 

    let (extractor:CellsExtractorBuilder) = new ExtractMonadBuilder<Wordprocessing.TableCell []>()

    type Extractor<'a> = ExtractMonad<'a, Wordprocessing.TableCell []> 


    let internal cellConsume = makeConsumeModule ()

    let getItem : Extractor<Wordprocessing.TableCell> = cellConsume.GetItem

    let getItems : int -> Extractor<Wordprocessing.TableCell []> = cellConsume.GetItems

    let position : Extractor<int> = cellConsume.Position

    let getInput : Extractor<Wordprocessing.TableCell []> = cellConsume.GetInput

    let inputCount : Extractor<int> = cellConsume.InputCount

    let satisfy : (Wordprocessing.TableCell -> bool) -> Extractor<Wordprocessing.TableCell> = cellConsume.Satisfy