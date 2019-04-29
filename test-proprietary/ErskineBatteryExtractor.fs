// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


module ErskineBatteryExtractor

open System.IO

open FSharp.Data

open DocSoup

let extractSiteDetails : TableExtractor< {| Name: string; SAI: string |} > = 
    mreturn {| Name = ""; SAI = "" |}