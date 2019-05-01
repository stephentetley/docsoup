// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace Extractors

module ErskineBatteryForm = 

    open System.IO

    open FSharp.Data

    open DocSoup

    let extractSiteDetails : Body.Extractor< {| Name: string; SAI: string; Outstation: string |} > = 
        Body.findTable (Table.innerTextIsMatch "Site Details") 
            &>> pipeM3 (Table.cell (1,1) &>> Cell.spacedText)
                       (Table.cell (2,1) &>> Cell.spacedText)
                       (Table.cell (4,1) &>> Cell.spacedText)
                       (fun name sai outs -> {| Name = name
                                              ; SAI = sai
                                              ; Outstation = outs |})


                                          
    let extractWorkDetails : Body.Extractor< {| Name: string; Date: string |} > = 
        Body.findTable (Table.firstCell &>> Cell.innerTextIsMatch "Testing & Recording of Site Work")
            &>> pipeM2 (Table.cell (2,1) &>> Cell.spacedText)
                        (Table.cell (3,1) &>> Cell.spacedText)
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

    let siteWorksExtractor : Document.Extractor<SiteWorksRow> = 
        Document.body 
            &>> pipeM2 extractSiteDetails 
                        extractWorkDetails
                        ( fun r1 r2 -> 
                            SiteWorksRow (siteName = r1.Name
                                        , engineerName = r2.Name
                                        , dateOfVisit = r2.Date
                                        , saiNumber = r1.SAI
                                        , outstationName = r1.Outstation ))

    let processSiteWorks (filePath:string) : Answer<SiteWorksRow>  =
        Document.runExtractor filePath siteWorksExtractor
