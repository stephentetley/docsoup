// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

module Mk5ReplacementExtract

open System.IO

open FSharp.Data

open DocSoup

let extractSiteInfo : Body.Extractor< {| Name: string; SAI: string |} > = 
    Body.findTable (Table.firstRow  &>> Row.isMatch [| "Site Information" |]) 
        &>> pipeM2 (Table.findNameValue1Row "Site Name")
                   (Table.findNameValue1Row "SAI Number")
                   (fun name sai -> {| Name = name; SAI = sai |})


let extractVisitInfo : Body.Extractor< {| Company: string; Name: string; Date: string |} > = 
    Body.findTable (Table.firstRow &>> Row.isMatch [| "Install Information" |]) 
        &>> pipeM3 (Table.cell (1,1) &>> Cell.paragraphsText)
                   (Table.cell (2,1) &>> Cell.paragraphsText)
                   (Table.cell (3,1) &>> Cell.paragraphsText)
                   (fun company eng date -> {| Company = company
                                             ; Name = eng
                                             ; Date = date |})



let extractOutstationInfo : Body.Extractor< {| Name: string; Address: string
                                            ; OldType: string; OldId: string
                                            ; NewType: string; NewId:string |} > = 
    Body.findTable (Table.firstRow &>> Row.isMatch [| "Outstation Information" |]) 
        &>> pipeM6 (Table.cell (1,1) &>> Cell.paragraphsText)
                   (Table.cell (2,1) &>> Cell.paragraphsText)
                   (Table.cell (3,1) &>> Cell.paragraphsText)
                   (Table.cell (4,1) &>> Cell.paragraphsText)
                   (Table.cell (5,1) &>> Cell.paragraphsText)
                   (Table.cell (6,1) &>> Cell.paragraphsText)
                   (fun name addr otype oid ntype nid -> 
                        {| Name = name
                        ; Address = addr
                        ; OldType = otype
                        ; OldId = oid
                        ; NewType = ntype
                        ; NewId = nid |})


let extractAdditionalComments : Body.Extractor< {| Issues: string
                                                ; QA: string
                                                ; Maintenance: string |} > = 
    Body.findTable (Table.firstRow &>> Row.isMatch [| "Additional Comment" |]) 
        &>> pipeM3 (Table.cell (1,1) &>> Cell.paragraphsText)
                   (Table.cell (2,1) &>> Cell.paragraphsText)
                   (Table.cell (3,1) &>> Cell.paragraphsText)
                   (fun issues qa mainten -> {| Issues = issues
                                             ; QA = qa
                                             ; Maintenance = mainten |})



[<Literal>]
let OutputSchema = 
    "Site Name(string), Sai Number(string), " +
    "Company(string), Engineer Name(string), Date Of Visit(string), " +
    "Outstation Name(string), Os Address(string), "+ 
    "Old Os Type(string), Old Os Serial Number(string), " +
    "New Os Type(string), New Os Serial Number(string), " +
    "Installation Issues(string), Quality and Performance(string), " +
    "Maintenance(string)"

type Mk5InstallTable = 
    CsvProvider< Sample = OutputSchema,
                 Schema = OutputSchema,
                 HasHeaders = true >

type Mk5InstallRow = Mk5InstallTable.Row


let mk5InstallExtractor : Document.Extractor<Mk5InstallRow> = 
    Document.body 
        &>> pipeM4 extractSiteInfo 
                    extractVisitInfo
                    extractOutstationInfo
                    extractAdditionalComments
                    ( fun r1 r2 r3 r4 -> 
                        Mk5InstallRow   ( siteName = r1.Name
                                        , saiNumber = r1.SAI
                                        , company = r2.Company
                                        , engineerName = r2.Name
                                        , dateOfVisit = r2.Date
                                        , outstationName = r3.Name
                                        , osAddress = r3.Address
                                        , oldOsType = r3.OldType
                                        , oldOsSerialNumber = r3.OldId
                                        , newOsType = r3.NewType
                                        , newOsSerialNumber = r3.NewId
                                        , installationIssues = r4.Issues
                                        , qualityAndPerformance = r4.QA
                                        , maintenance = r4.Maintenance
                                        ))

let processMk5Install (filePath:string) : Answer<Mk5InstallRow>  =
    Document.runExtractor filePath mk5InstallExtractor



