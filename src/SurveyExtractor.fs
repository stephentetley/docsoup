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

///// Returns "" if no cell matches the search.
//let getFieldValue (search:string) (matchCase:bool) : Table1<string> = 
//    let good = 
//        withCellM (findCell search matchCase &>>= cellRight) <| getCellText
//    good <|||> t1return ""

///// Returns "" if no cell matches the search.
///// This speeds things up a bit, but for our use case here we are
///// concerned about layout changes.
//let getFieldValueByRow (row:int) : Table1<string> = 
//    let good = 
//        withCellM (getCellByIndex row 2) <| getCellText
//    good <|||> t1return ""


///// Returns "" if no cell matches the search.
//let getFieldValuePattern (search:string) : Table1<string> = 
//    let good = 
//        withCellM (findCellByPattern search &>>= cellRight) <| getCellText
//    good <|||> t1return ""


//let assertHeaderCell (str:string) : RowExtractor<unit> = 
//    withCellM (getCellByIndex 1 1) <| assertCellText str

    
//let assertHeaderCellPattern (pattern:string) : RowExtractor<unit> = 
//    withCellM (getCellByIndex 1 1) <| assertCellMatches pattern

//let ignoreTable (tableHeader:string) : TablesExtractor<unit> = 
//    parseTable (assertHeaderCell tableHeader)

//let tableNot (tableHeader:string) : TablesExtractor<unit> = 
//    parseTable (assertCellTextNot tableHeader)


// *************************************
// Survey parsers

let extractSiteDetails : RowParser<SiteInfo> = 
    parseRows { 
        do! row (assertCellText "Site Details")
        let! sname          = row (assertCellText "Site Name" &>>>. cellText)
        let! uid            = row (assertCellText "SAI Number" &>>>. cellText)
        let! discharge      = row (assertCellText "Discharge Name" &>>>. cellText)
        let! watercourse    = row (assertCellText "Receiving Watercourse" &>>>. cellText)
        return {
            SiteName = sname
            SaiNumber = uid
            DischargeName = discharge
            ReceivingWatercourse = watercourse
            }
    }


let extractSurveyInfo : RowParser<SurveyInfo> = 
    table1 { 
        do! assertHeaderCell "Survey Information"
        let! name       = getFieldValue "Engineer Name" false
        let! sdate      = getFieldValue "Date of Survey" false
        return { 
            EngineerName = name
            SurveyDate = sdate
        }
    }


let extractOutstationInfo : Table1<OutstationInfo> = 
    table1 { 
        do! assertHeaderCell "RTU Outstation"
        let! name       = getFieldValue "Outstation Name" false
        let! rtuAddr    = getFieldValue "RTU Address" false
        let! otype      = getFieldValue "Outstation Type" false
        let! snumber    = getFieldValue "Serial Number" false
        return { 
            OutstationName = name
            RtuAddress = rtuAddr
            OutstationType = otype 
            SerialNumber = snumber } 
    }

 
/// Focus should already be limited to the table in question.
let extractRelay (relayNumber:int) : Table1<RelaySetting> = 
    let funPattern  = sprintf "Relay*%i*Function" relayNumber
    let onPattern   = sprintf "Relay*%i*On" relayNumber
    let offPattern  = sprintf "Relay*%i*Off" relayNumber
    table1 { 
        let! relayfunction  = getFieldValuePattern funPattern
        let! onSetPt        = getFieldValuePattern onPattern
        let! offSetPt       = getFieldValuePattern offPattern
        return { 
            RelayNumber = relayNumber
            RelayFunction = relayfunction
            OnSetPoint = onSetPt
            OffSetPoint = offSetPt 
            }
        }


let extractRelays : Table1<RelaySetting list> = 
    DocSoup.TableExtractor1.mapM extractRelay [1..6] 
        &|>>> List.filter (fun (r:RelaySetting) -> not r.isEmpty)

let extractUltrasonicMonitorInfo : Table1<UltrasonicMonitorInfo> = 
    table1 {
        do! assertHeaderCell "Ultrasonic Level Control"
        let! disName    = getFieldValue "Discharge Being Monitored" false
        let! procName   = getFieldValueByRow 3
        let! manuf      = getFieldValue "Manufacturer" false
        let! model      = getFieldValue "Model" false
        let! snumber    = getFieldValue "Serial Number" false
        let! piTag      = getFieldValue "P & I Tag" false
        let! emptyDist  = getFieldValue "Empty Distance" false
        let! span       = getFieldValue "Span" false
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

let extractUltrasonicSensorInfo : Table1<UltrasonicSensorInfo> = 
    table1 {
        do! assertHeaderCell "Ultrasonic Sensor Head"
        let! manuf      = getFieldValue "Manufacturer" false
        let! model      = getFieldValue "Model" false
        let! snumber    = getFieldValue "Serial Number" false
        let! location   = getFieldValuePattern "Location of Sensor"
        let! gridRef    = getFieldValuePattern "Grid Ref"
        return { 
            SensorManufacturer = manuf
            SensorModel = model
            SerialNumber = snumber 
            LocationOfSensor = location
            GridRef = gridRef }
    } <&??> "extractUltrasonicSensorInfo"

let extractUltrasonicInfo1 : TablesExtractor<UltrasonicInfo> = 
    parseTables { 
        let! monitor = parseTable extractUltrasonicMonitorInfo <&?> "ultrasonic failed"
        let! sensor  = parseTable extractUltrasonicSensorInfo
        return { 
            MonitorInfo = monitor
            SensorInfo = sensor } 
    }

/// To ignore
let tableUltrasonicPhotos : Table1<unit> = 
    assertHeaderCell "Ultrasonic Photos" 

let extractOverflowType : Table1<OverflowType> =  
    table1 { 
        let! a = TableExtractor1.optional <| findCell "Screen to Invert" false
        let! b = TableExtractor1.optional <| findCell "Emergency overflow level" false
        match a,b with
        | Some _, Some _ -> return SCREENED
        | _, _ -> return UNSCREENED
    }


let extractOverflowChamberInfo : Table1<OverflowChamberInfo> =  
    table1 { 
        do! assertHeaderCell "Overflow Chamber" 
        let! disName        = getFieldValue "Discharge Name" false
        let! chamberName    = getFieldValuePattern "Name of Chamber"
        let! gridRef        = getFieldValue "Grid Ref" false
        let! screened       = getFieldValuePattern "Overflow Screened"
        return { 
            DischargeName = disName
            ChamberName = chamberName
            OverflowGridRef = gridRef
            IsScreened = screened }
    }

let extractOverflowChamberMetrics : Table1<OverflowChamberMetrics> =  
    table1 { 
        do! assertHeaderCell "Chamber Measurements"
        let! otype      = extractOverflowType
        let! name       = getFieldValue "Chamber Name" false
        let! roofDist   = getFieldValue "Roof Slab to Invert" false
        let! usDist     = getFieldValue "Transducer Face to Invert" false
        let! ovDist     = getFieldValue "Overflow level to Invert" false
        let! scDist     = getFieldValuePattern "Bottom*Screen*Invert"
        let! emDist     = getFieldValuePattern "Emergency*Invert"
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
let tableOverflowPhoto : Table1<unit> = 
    assertHeaderCell "Overflow Chamber Photo"
    

let extractOutfallInfo : Table1<OutfallInfo> = 
    table1 { 
        do! assertHeaderCell "Outfall"
        let! dname      = getFieldValue "Discharge Name" false
        let! gridRef    = getFieldValue "Grid Ref" false
        let! proven     = getFieldValue "Outfall Proven" false
        return { 
            DischargeName = dname
            OutfallGridRef = gridRef
            OutfallProven = proven
        }
    }

/// To ignore
let tableOutfallPhoto : Table1<unit> = 
    table1 {
        do! assertHeaderCell "Outfall Photo"
        return ()
    }

let extractScopeOfWorks : Table1<string> = 
    table1 {
        do! assertHeaderCell "Scope of Works"
        let! ans = withCellM (getCellByIndex 3 1) <| getCellText
        return ans
    }

/// Single table - Title (1,1) = "Appendix", data in cell (r2,c1):
let extractAppendix : Table1<string> =  
    assertHeaderCellPattern "Appendix" &>>>. 
        (withCellM (getCellByIndex 2 1) <| getCellText)


let ultrasonicTable : TablesExtractor<UltrasonicInfo option> = 
    let info = extractUltrasonicInfo1 |>>> Some
    let photos = parseTable tableUltrasonicPhotos |>>> (fun () -> None)
    (photos <||> info) <&?> "ultrasonicTable"


let sectionUltrasonics : TablesExtractor<UltrasonicInfo list>= 
    many1 ultrasonicTable |>>> (List.choose id)
    

/// list<Overflow Chamber> * list<Chamber Measurements>
/// ignore <Overflow Chamber Photo>

type ChamberTable = 
    | InfoTable of OverflowChamberInfo
    | MetricsTable of OverflowChamberMetrics
    | PhotoTable of unit



let chamberTable : Table1<ChamberTable> = 
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


let outfallTable : Table1<OutfallInfo option> = 
    let info = extractOutfallInfo &|>>> Some
    let photo = tableOutfallPhoto &|>>> (fun () -> None)
    swapTableError "outfallTable" (photo <|||> info)


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
