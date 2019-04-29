// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


module PBCommissioningForm

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


let extractCommissionItem (fileName:string) : TableExtractor<CommissionItem> = 
    tableExtractor { 
        let! (visit,crew)                   = row 1 &>> tupleM2 (cell 1 &>> cellParagraphsText) (cell 3 &>> cellParagraphsText)
        let! rtuName                        = row 2 &>> cell 1 &>> cellParagraphsText
        let! manholeLoc                     = row 3 &>> cell 1 &>> cellParagraphsText
        let! accessNotes                    = row 4 &>> cell 1 &>> cellParagraphsText
        let! device1                        = row 8 &>> cell 1 &>> cellParagraphsText
        let! deviceSerialNo                 = row 9 &>> cell 1 &>> cellParagraphsText
        let! batterySerialNo                = row 11 &>> cell 1 &>> cellParagraphsText
        let! sensorType                     = row 13 &>> cell 1 &>> cellParagraphsText
        let! sensorNote                     = row 14 &>> cell 1 &>> cellParagraphsText
        let! slabToInvert                   = row 25 &>> cell 1 &>> cellParagraphsText
        let! offset                         = row 26 &>> cell 1 &>> cellParagraphsText
        let! liquidLevel                    = row 27 &>> cell 1 &>> cellParagraphsText
        let! overflowToInvert               = row 28 &>> cell 1 &>> cellParagraphsText
        let! emergencyOverflowToInvert      = row 29 &>> cell 1 &>> cellParagraphsText
        let! blankingDistance               = row 30 &>> cell 1 &>> cellParagraphsText
        let! span                           = row 31 &>> cell 1 &>> cellParagraphsText
        let! abortedVisitYesNo              = row 33 &>> cell 1 &>> cellParagraphsText
        let! visitInfo                      = row 34 &>> cell 1 &>> cellParagraphsText
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

let process1 (fileName:string) : DocumentExtractor<CommissionItem> = 
    body &>> table 0 &>> extractCommissionItem fileName

let processCommissionForm(filePath:string) : Answer<CommissionItem>  =
    let name = FileInfo(filePath).Name
    runDocumentExtractor filePath (process1 name)
    


