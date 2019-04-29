// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

module Mk5ReplacementExtract

open System.IO

open FSharp.Data

open DocSoup

let extractSiteInfo : BodyExtractor< {| Name: string; SAI: string |} > = 
    findTable (tableFirstRow &>> rowIsMatch [| "Site Information" |]) 
        &>> pipeM2 (tableCell 1 1 &>> cellParagraphsText)
                   (tableCell 2 1 &>> cellParagraphsText)
                   (fun name sai -> {| Name = name; SAI = sai |})


let extractInstallInfo : BodyExtractor< {| Company: string; Name: string; Date: string |} > = 
    findTable (tableFirstRow &>> rowIsMatch [| "Install Information" |]) 
        &>> pipeM3 (tableCell 1 1 &>> cellParagraphsText)
                   (tableCell 2 1 &>> cellParagraphsText)
                   (tableCell 3 1 &>> cellParagraphsText)
                   (fun company eng date -> {| Company = company
                                             ; Name = eng
                                             ; Date = date |})

