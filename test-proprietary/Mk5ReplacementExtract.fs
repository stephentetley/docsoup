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


let extractVisitInfo : BodyExtractor< {| Company: string; Name: string; Date: string |} > = 
    findTable (tableFirstRow &>> rowIsMatch [| "Install Information" |]) 
        &>> pipeM3 (tableCell 1 1 &>> cellParagraphsText)
                   (tableCell 2 1 &>> cellParagraphsText)
                   (tableCell 3 1 &>> cellParagraphsText)
                   (fun company eng date -> {| Company = company
                                             ; Name = eng
                                             ; Date = date |})



let extractOutstationInfo : BodyExtractor< {| Name: string; Address: string
                                            ; OldType: string; OldId: string
                                            ; NewType: string; NewId:string |} > = 
    findTable (tableFirstRow &>> rowIsMatch [| "Outstation Information" |]) 
        &>> pipeM6 (tableCell 1 1 &>> cellParagraphsText)
                   (tableCell 2 1 &>> cellParagraphsText)
                   (tableCell 3 1 &>> cellParagraphsText)
                   (tableCell 4 1 &>> cellParagraphsText)
                   (tableCell 5 1 &>> cellParagraphsText)
                   (tableCell 6 1 &>> cellParagraphsText)
                   (fun name addr otype oid ntype nid -> 
                        {| Name = name
                        ; Address = addr
                        ; OldType = otype
                        ; OldId = oid
                        ; NewType = ntype
                        ; NewId = nid |})


let extractAdditionalComments : BodyExtractor< {| Issues: string
                                                ; QA: string
                                                ; Maintenance: string |} > = 
    findTable (tableFirstRow &>> rowIsMatch [| "Additional Comment" |]) 
        &>> pipeM3 (tableCell 1 1 &>> cellParagraphsText)
                   (tableCell 2 1 &>> cellParagraphsText)
                   (tableCell 3 1 &>> cellParagraphsText)
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


let mk5InstallExtractor : DocumentExtractor<Mk5InstallRow> = 
    body &>> pipeM4 extractSiteInfo 
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
    runDocumentExtractor filePath mk5InstallExtractor



