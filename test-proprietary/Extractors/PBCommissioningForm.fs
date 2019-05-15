// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace Extractors

module PBCommissioningForm = 

    open System.IO

    open FSharp.Data

    open DocSoup


    [<Literal>]
    let CommissionFormSchema = 
        "File Name(string), Date Of Visit(string), Crew(string)," +
        "RTU Name(string), Manhole Location(string), Access Notes(string)," +
        "Device1(string), Device Serial No(string), Battery Serial No(string)," +
        "Sensor Type(string), Sensor Note(string)," + 
        "Slab To Invert(string), Offset(string), Liquid Level(string)," +
        "Overflow To Invert(string), Emergency Overflow To Invert(string)," +
        "Blanking Distance(string), Span(string)," +
        "Aborted Visit Yes No(string), Visit Info(string)"


    /// Trick - setting Sample to ExportSchema rather than a sample "row" writes the schema as
    /// Headers in the output.
    type CommissionsTable = 
        CsvProvider< Sample = CommissionFormSchema,
                     Schema = CommissionFormSchema,
                     HasHeaders = true >

    type CommissionItem = CommissionsTable.Row


    let extractCommissionItem (fileName:string) : Table.Extractor<CommissionItem> = 
        Table.extractor { 
            let! (visit,crew)                   = Table.row 1 &>> tupleM2 (Row.cell 1 &>> Cell.spacedText) (Row.cell 3 &>> Cell.spacedText)
            let! rtuName                        = Table.cell (2,1) &>> Cell.spacedText
            let! manholeLoc                     = Table.cell (3,1) &>> Cell.spacedText
            let! accessNotes                    = Table.cell (4,1) &>> Cell.spacedText
            let! device1                        = Table.cell (8,1) &>> Cell.spacedText
            let! deviceSerialNo                 = Table.cell (9,1) &>> Cell.spacedText
            let! batterySerialNo                = Table.cell (11,1) &>> Cell.spacedText
            let! sensorType                     = Table.cell (13,1) &>> Cell.spacedText
            let! sensorNote                     = Table.cell (14,1) &>> Cell.spacedText
            let! slabToInvert                   = Table.cell (25,1) &>> Cell.spacedText
            let! offset                         = Table.cell (26,1) &>> Cell.spacedText
            let! liquidLevel                    = Table.cell (27,1) &>> Cell.spacedText
            let! overflowToInvert               = Table.cell (28,1) &>> Cell.spacedText
            let! emergencyOverflowToInvert      = Table.cell (29,1) &>> Cell.spacedText
            let! blankingDistance               = Table.cell (30,1) &>> Cell.spacedText
            let! span                           = Table.cell (31,1) &>> Cell.spacedText
            let! abortedVisitYesNo              = Table.cell (33,1) &>> Cell.spacedText
            let! visitInfo                      = Table.cell (34,1) &>> Cell.spacedText
            return (new CommissionItem ( fileName = fileName
                                      , dateOfVisit = visit
                                      , crew = crew
                                      , rtuName = rtuName
                                      , manholeLocation = manholeLoc
                                      , accessNotes = accessNotes
                                      , device1 = device1
                                      , deviceSerialNo = deviceSerialNo
                                      , batterySerialNo = batterySerialNo
                                      , sensorType = sensorType
                                      , sensorNote = sensorNote
                                      , slabToInvert = slabToInvert
                                      , offset = offset
                                      , liquidLevel = liquidLevel
                                      , overflowToInvert = overflowToInvert
                                      , emergencyOverflowToInvert = emergencyOverflowToInvert
                                      , blankingDistance = blankingDistance
                                      , span = span
                                      , abortedVisitYesNo = abortedVisitYesNo
                                      , visitInfo = visitInfo
                                      ))
        }

    let process1 (fileName:string) : Document.Extractor<CommissionItem> = 
        Document.body &>> Body.table 0 &>> extractCommissionItem fileName

    let processCommissionForm(filePath:string) : Result<CommissionItem, ErrMsg>  =
        let name = FileInfo(filePath).Name
        Document.runExtractor filePath (process1 name)
    


