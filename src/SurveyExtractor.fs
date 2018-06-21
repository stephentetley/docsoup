// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module SurveyExtractor

open DocSoup.Base
open DocSoup.RowExtractor
open DocSoup.TablesExtractor
open DocSoup

open SurveySyntax


// Note - layout changes are expected for the input documents we are querying here.
// This makes us favour by-name access rather than (faster) by-index access.



// *************************************
// Helpers

let sw (msg:string) (ma:TablesExtractor<'a>) : TablesExtractor<'a> =
    parseTables { 
        let stopWatch = System.Diagnostics.Stopwatch.StartNew()
        do printfn "%s" msg
        let! ans = ma
        do stopWatch.Stop()
        do printfn "... time(ms) %d" stopWatch.ElapsedMilliseconds
        return ans
        }



// *************************************
// Utility parsers

let rowOf1 (exactMatch:string) : RowParser<unit> = 
    row (assertCellText exactMatch &>>>. endOfRow)

let rowOf1Matching (patternMatch:string) : RowParser<unit> = 
    row (assertCellWordMatch patternMatch &>>>. endOfRow)

let rowOf2 (exactMatch:string) : RowParser<string> = 
    let errMsg = sprintf "rowOf2 failed - looking for '%s'" exactMatch
    row (assertCellText exactMatch &>>>. cellText) <&??> errMsg

let rowOf2Matching (patternMatch:string) : RowParser<string> = 
    row (assertCellWordMatch patternMatch &>>>. cellText)


let tableNot (tableHeader:string) : TablesExtractor<unit> = 
    parseTable (row <| assertCellTextNot tableHeader)


// *************************************
// Survey parsers

let extractSiteDetails : RowParser<SiteInfo> = 
    parseRows { 
        do! rowOf1 "Site Details"
        let! sname          = rowOf2 "Site Name"
        let! uid            = rowOf2 "SAI Number"
        let! discharge      = skipRowsTill <| rowOf2 "Discharge Name"
        let! watercourse    = skipRowsTill <| rowOf2 "Receiving Watercourse"
        return {
            SiteName = sname
            SaiNumber = uid
            DischargeName = discharge
            ReceivingWatercourse = watercourse
            }
    }


let extractSurveyInfo : RowParser<SurveyInfo> = 
    parseRows { 
        do! rowOf1 "Survey Information"
        let! name       = skipRowsTill <| rowOf2Matching "*Engineer Name"
        let! sdate      = rowOf2 "Date of Survey"
        return { 
            EngineerName = name
            SurveyDate = sdate
        }
    }


let extractOutstationInfo : RowParser<OutstationInfo> = 
    parseRows { 
        do! rowOf1 "RTU Outstation"
        let! name       = 
            skipRowsTill <| rowOf2 "RTU Outstation Name"
        let! rtuAddr    = rowOf2 "RTU Address"
        let! otype      = rowOf2 "Outstation Type" 
        let! snumber    = rowOf2 "Outstation Serial Number" 
        return { 
            OutstationName = name
            RtuAddress = rtuAddr
            OutstationType = otype 
            SerialNumber = snumber } 
    }

 
/// Focus should already be limited to the table in question.
let extractRelay (relayNumber:int) : RowParser<RelaySetting> = 
    parseRows { 
        let! relayfunction  = skipRowsTill <| rowOf2Matching "Relay*Function"
        let! onSetPt        = rowOf2Matching "Relay*On*"
        let! offSetPt       = rowOf2Matching "Relay*Off*"
        return { 
            RelayNumber = relayNumber
            RelayFunction = relayfunction
            OnSetPoint = onSetPt
            OffSetPoint = offSetPt 
            }
        }


let extractRelays : RowParser<RelaySetting list> = 
    DocSoup.RowExtractor.mapM extractRelay [1..6] 
        &|>>> List.filter (fun (r:RelaySetting) -> not r.isEmpty)

let extractUltrasonicMonitorInfo : RowParser<UltrasonicMonitorInfo> = 
    parseRows {
        do! rowOf1 "Ultrasonic Level Control"
        let! disName    = rowOf2 "Discharge Being Monitored"
        let! procName   = rowOf2Matching "Name of Process*"
        let! manuf      = rowOf2 "Manufacturer"
        let! model      = rowOf2 "Model"
        let! snumber    = rowOf2 "Serial Number"
        let! piTag      = rowOf2Matching "* Tag"
        let! emptyDist  = skipRowsTill <| rowOf2 "Empty Distance"
        let! span       = rowOf2 "Span"
        let! relays     = extractRelays
        return { 
            MonitoredDischarge = disName
            ProcessOrFacilityName = procName
            MonitorManufacturer = manuf
            MonitorModel = model
            SerialNumber = snumber
            PITag = piTag
            EmptyDistance = emptyDist
            Span = span
            Relays = relays }
    }  <&??> "extractUltrasonicMonitorInfo"

let extractUltrasonicSensorInfo : RowParser<UltrasonicSensorInfo> = 
    parseRows {
        do! rowOf1 "Ultrasonic Sensor Head"
        let! manuf      = rowOf2 "Manufacturer"
        let! model      = rowOf2 "Model"
        let! snumber    = rowOf2 "Serial Number"
        let! location   = skipRowsTill <| rowOf2Matching "Location of Sensor*"
        let! gridRef    = rowOf2Matching "Grid Ref*"
        return { 
            SensorManufacturer = manuf
            SensorModel = model
            SerialNumber = snumber 
            LocationOfSensor = location
            GridRef = gridRef }
    } <&??> "extractUltrasonicSensorInfo"

let extractUltrasonicInfo1 : TablesExtractor<UltrasonicInfo> = 
    parseTables { 
        let! monitor = parseTable extractUltrasonicMonitorInfo
        let! sensor  = parseTable extractUltrasonicSensorInfo
        return { 
            MonitorInfo = monitor
            SensorInfo = sensor } 
    }

/// To ignore
let tableUltrasonicPhotos : RowParser<unit> = 
    row (assertCellText "Ultrasonic Photos")


let extractOverflowChamberInfo : RowParser<OverflowChamberInfo> =  
    parseRows { 
        do! rowOf1 "Overflow Chamber"                  
        let! disName        = rowOf2 "Discharge Name"
        let! chamberName    = rowOf2Matching "Name of Chamber*"
        let! gridRef        = rowOf2 "Grid Ref"
        let! screened       = rowOf2Matching "Is Overflow Screened*"
        return { 
            DischargeName = disName
            ChamberName = chamberName
            OverflowGridRef = gridRef
            IsScreened = screened }
    }



let extractOverflowType : RowParser<OverflowType> =  
    let proc : RowParser<string * string> = 
        parseRows { 
            let! a = skipRowsTill <| rowOf2Matching "*Screen to Invert"
            let! b = skipRowsTill <| rowOf2Matching "*Emergency overflow level*"
            return (a,b)
        }
    let answer (ans:option<string * string>) : OverflowType = 
        match ans with
        | Some(_,_) -> SCREENED
        | None -> UNSCREENED

    RowExtractor.lookahead (RowExtractor.optional proc &|>>> answer)
    


let extractOverflowChamberMetrics : RowParser<OverflowChamberMetrics> =  
    parseRows { 
        do! rowOf1 "Chamber Measurements"
        let! otype      = extractOverflowType
        printfn "___ Chamber Measurements otype=%A" otype
        do! printIx ()
        let! name       = fatal "chamber name" <| rowOf2 "Chamber Name"
        do! printIx ()
        printfn "___ 0.5"
        let! roofDist   = rowOf2Matching "*Roof Slab to Invert"
        printfn "___ 1"
        let! usDist     = rowOf2Matching "*Transducer Face to Invert"
        printfn "___ 2"
        let! scDist     = 
            rowOf2Matching "*Bottom of Screen to Invert" <|||> rereturn ""
        printfn "___ 3"    
        let! ovDist     = rowOf2Matching  "*Overflow level to Invert"
        printfn "___ 4"
        let! emDist     = 
            rowOf2Matching "*Emergency*to Invert" <|||> rereturn ""
            
        printfn "___ Chamber Measurements DONE"
        return { 
            OverflowType = otype
            ChamberName = name
            RoofToInvert = roofDist
            UsFaceToInvert = usDist 
            OverflowToInvert = ovDist
            ScreenToInvert = scDist
            EmergencyOverflowToInvert = emDist 
        }
    }

/// To ignore
let tableOverflowPhoto : RowParser<unit> = 
    row (assertCellText "Overflow Chamber Photo")
    

let extractOutfallInfo : RowParser<OutfallInfo> = 
    parseRows { 
        do! rowOf1 "Outfall"
        printfn "___ Outfall"
        let! dname      = rowOf2 "Discharge Name"
        let! gridRef    = rowOf2Matching "Grid Ref*"
        let! proven     = rowOf2Matching "Outfall Proven*"
        printfn "___ Outfall DONE"
        return { 
            DischargeName = dname
            OutfallGridRef = gridRef
            OutfallProven = proven
        }
    }

/// To ignore
let tableOutfallPhoto : RowParser<unit> = 
    row (assertCellText "Outfall Photo")

let extractScopeOfWorks : RowParser<string> = 
    parseRows {
        do! rowOf1 "Scope of Works"
        do! skipRow
        let! ans = row <| cellText
        return ans
    }

/// Single table - Title (1,1) = "Appendix", data in cell (r2,c1):
let extractAppendix : RowParser<string> =  
    parseRows {
        do! rowOf1Matching "Appendix *"
        let! ans = row <| cellText
        return ans
    }
let ultrasonicTable : TablesExtractor<UltrasonicInfo option> = 
    let photos = parseTable tableUltrasonicPhotos |>>> (fun () -> None)
    let info = extractUltrasonicInfo1 |>>> Some
    (photos <||> info) <&?> "ultrasonicTable"


let sectionUltrasonics : TablesExtractor<UltrasonicInfo list>= 
    many1 ultrasonicTable |>>> (List.choose id)
    

/// list<Overflow Chamber> * list<Chamber Measurements>
/// ignore <Overflow Chamber Photo>

type ChamberTable = 
    | InfoTable of OverflowChamberInfo
    | MetricsTable of OverflowChamberMetrics
    | PhotoTable of unit



let chamberTable : RowParser<ChamberTable> = 
    let info    = extractOverflowChamberInfo &|>>> InfoTable
    let metrics = extractOverflowChamberMetrics &|>>> MetricsTable 
    let photo   = tableOverflowPhoto &|>>> PhotoTable
    (metrics <|||> info <|||> photo) 

let private getOverflowLists (source: ChamberTable list) 
                                : OverflowChamberInfo list * OverflowChamberMetrics list = 
    let rec work ac bc xs = 
        match xs with
        | [] -> (List.rev ac, List.rev bc)
        | InfoTable a :: rest -> work (a::ac) bc rest
        | MetricsTable b :: rest -> work ac (b::bc) rest
        | PhotoTable _ :: rest -> work ac bc rest
    work [] [] source
            
let sectionChambers : TablesExtractor<OverflowChamberInfo list * OverflowChamberMetrics list>= 
    many1 (parseTable chamberTable) |>>> getOverflowLists


let outfallTable : RowParser<OutfallInfo option> = 
    let info = extractOutfallInfo &|>>> Some
    let photo = tableOutfallPhoto &|>>> (fun () -> None)
    swapRowError "outfallTable" (photo <|||> info)


/// Note table parser would find finds "OutFall Photos" if we just looked for 
/// "Outfall".
let sectionOutfalls : TablesExtractor<OutfallInfo list>= 
    many1 (parseTable outfallTable) |>>> (List.choose id)
        





let parseSurvey : TablesExtractor<Survey> = 
    parseTables { 
        let! site           = 
            sw "site"           <| parseTable extractSiteDetails      

        let! surveyInfo     = 
            sw "survey info"    <| parseTable extractSurveyInfo

        do! skipMany (tableNot "RTU Outstation")
        let! outstation     = 
            sw "outstation"     <| parseTable  extractOutstationInfo

        do! skipMany (tableNot "Ultrasonic Level Control")
        let! ultrasonics    = sw "ultrasonics"      sectionUltrasonics

        let! (chambers, metrics)    = 
            sw "chambers"         sectionChambers
        let! outfalls       = sw "outfalls"         sectionOutfalls
        let! scopeText      = 
            sw "scope-of-works"  <| parseTable   extractScopeOfWorks
        let! appendixText   = 
            sw "appendix"        <| parseTable  extractAppendix
        return { 
            SiteDetails = site
            SurveyInfo = surveyInfo
            OutstationInfo = outstation
            UltrasonicInfos = ultrasonics
            ChamberMetrics = metrics
            ChamberInfos = chambers
            OutfallInfos = outfalls
            ScopeOfWorks = scopeText
            AppendixText = appendixText
            }
    }
