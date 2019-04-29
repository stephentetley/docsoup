// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


module ErskineBatteryExtract

open System.IO

open FSharp.Data

open DocSoup

let extractSiteDetails 
        : BodyExtractor< {| Name: string; SAI: string; Outstation: string |} > = 
    findTable (tableInnerTextMatch "Site Details") 
        &>> pipeM3 (tableCell 1 1 &>> cellInnerText)
                   (tableCell 2 1 &>> cellInnerText)
                   (tableCell 4 1 &>> cellInnerText)
                   (fun name sai outs -> {| Name = name
                                          ; SAI = sai
                                          ; Outstation = outs |})


                                          
let extractWorkDetails 
        : BodyExtractor< {| Name: string; Date: string |} > = 
    findTable (tableCell 0 0 &>> cellInnerTextMatch "Testing & Recording of Site Work")
        &>> pipeM2 (tableCell 2 1 &>> cellInnerText)
                    (tableCell 3 1 &>> cellInnerText)
                    (fun name date -> {| Name = name
                                        ; Date = date |})

[<Literal>]
let OutputSchema = 
    "Site Name(string), Engineer Name(string), Date Of Visit(string), " +
    "Sai Number(string), Outstation Name(string)"

type SiteWorksTable = 
    CsvProvider< Sample = OutputSchema,
                 Schema = OutputSchema,
                 HasHeaders = true >

type SiteWorksRow = SiteWorksTable.Row

let siteWorksExtractor : DocumentExtractor<SiteWorksRow> = 
    body &>> pipeM2 extractSiteDetails 
                    extractWorkDetails
                    ( fun r1 r2 -> 
                        SiteWorksRow (siteName = r1.Name
                                    , engineerName = r2.Name
                                    , dateOfVisit = r2.Date
                                    , saiNumber = r1.SAI
                                    , outstationName = r1.Outstation ))

let processSiteWorks (filePath:string) : Answer<SiteWorksRow>  =
    runDocumentExtractor filePath siteWorksExtractor
