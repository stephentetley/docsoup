// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module SurveyExtractor

open DocSoup.Base
open DocSoup.TableExtractor1
open DocSoup.DocExtractor
open DocSoup

open SurveySyntax


// Note - layout changes are expected for the input documents we are querying here.
// This makes us favour by-name access rather than (faster) by-index access.



// *************************************
// Helpers

let sw (msg:string) (ma:DocExtractor<'a>) : DocExtractor<'a> =
    docExtract { 
        let stopWatch = System.Diagnostics.Stopwatch.StartNew()
        do printfn "%s" msg
        let! ans = ma
        do stopWatch.Stop()
        do printfn "... time(ms) %d" stopWatch.ElapsedMilliseconds
        return ans
        }



// *************************************
// Utility parsers

/// Returns "" if no cell matches the search.
let getFieldValue (search:string) (matchCase:bool) : TableExtractor1<string> = 
    let good = 
        withCellM (findCell search matchCase &>>= cellRight) <| getCellText
    good <|||> t1return ""

/// Returns "" if no cell matches the search.
/// This speeds things up a bit, but for our use case here we are
/// concerned about layout changes.
let getFieldValueByRow (row:int) : TableExtractor1<string> = 
    let good = 
        withCellM (getCellByIndex row 2) <| getCellText
    good <|||> t1return ""


/// Returns "" if no cell matches the search.
let getFieldValuePattern (search:string) : TableExtractor1<string> = 
    let good = 
        withCellM (findCellByPattern search &>>= cellRight) <| getCellText
    good <|||> t1return ""


let section (number:int) 
                (ma:DocExtractor<'a>) : DocExtractor<'a> = 
    let startMarker = sprintf "Section %i" number
    advanceM (findTextEnd startMarker true) >>>. ma

let assertHeaderCell (str:string) : TableExtractor1<unit> = 
    withCellM (getCellByIndex 1 1) <| assertCellText str

// *************************************
// Survey parsers

let extractSiteDetails : TableExtractor1<SiteInfo> = 
    table1 { 
        do! assertHeaderCell "Site Details"
        let! sname          = getFieldValue "Site Name" false
        let! uid            = getFieldValue "SAI Number" false
        let! discharge      = getFieldValue "Discharge Name" false
        let! watercourse    = getFieldValue "Receiving Watercourse" false
        return {
            SiteName = sname
            SaiNumber = uid
            DischargeName = discharge
            ReceivingWatercourse = watercourse
            }
    }


let extractSurveyInfo : TableExtractor1<SurveyInfo> = 
    table1 { 
        do! assertHeaderCell "Survey Information"
        let! name       = getFieldValue "Engineer Name" false
        let! sdate      = getFieldValue "Date of Survey" false
        return { 
            EngineerName = name
            SurveyDate = sdate
        }
    }


let extractOutstationInfo : TableExtractor1<OutstationInfo> = 
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
let extractRelay (relayNumber:int) : TableExtractor1<RelaySetting> = 
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


let extractRelays : TableExtractor1<RelaySetting list> = 
    DocSoup.TableExtractor1.mapM extractRelay [1..6] &|>>> List.filter (fun (r:RelaySetting) -> not r.isEmpty)

let extractUltrasonicMonitorInfo : TableExtractor1<UltrasonicMonitorInfo> = 
    table1 {
        do! assertHeaderCell "Ultrasonic Level Control"
        let! disName    = getFieldValue "Discharge Being Monitored" false
        let! procName   = getFieldValue "Process or Facility" false
        let! manuf      = getFieldValue "Manufacturer" false
        let! model      = getFieldValue "Model" false
        let! snumber    = getFieldValue "Serial Number" false
        let! piTag      = getFieldValue "P & I Tag" false
        let! emptyDist  = getFieldValue "Empty Distance" false
        let! span       = getFieldValue "Span" false
        let! relays     =  extractRelays
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
    } 

let extractUltrasonicSensorInfo : TableExtractor1<UltrasonicSensorInfo> = 
    table1 {
        do! assertHeaderCell "Ultrasonic Sensor Head"
        let! manuf      = getFieldValue "Manufacturer" false
        let! model      = getFieldValue "Model" false
        let! snumber    = getFieldValue "Serial Number" false
        return { 
            SensorManufacturer = manuf
            SensorModel = model
            SerialNumber = snumber 
            LocationOfSensor = ""
            GridRef = "" }
    }

let extractUltrasonicInfo1 : DocExtractor<UltrasonicInfo> = 
    docExtract { 
        let! monitor = nextTable extractUltrasonicMonitorInfo
        let! sensor  = nextTable extractUltrasonicSensorInfo
        return { 
            MonitorInfo = monitor
            SensorInfo = sensor } 
    }

/// To ignore
let tableUltrasonicPhotos : DocExtractor<unit> = 
    nextTable (assertHeaderCell "Ultrasonic Photos" )

let extractOverflowType : TableExtractor1<OverflowType> =  
    table1 { 
        let! a = TableExtractor1.optional <| findCell "Screen to Invert" false
        let! b = TableExtractor1.optional <| findCell "Emergency overflow level" false
        match a,b with
        | Some _, Some _ -> return SCREENED
        | _, _ -> return UNSCREENED
    }


let extractOverflowChamberInfo : TableExtractor1<OverflowChamberInfo> =  
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

let extractOverflowChamberMetrics : TableExtractor1<OverflowChamberMetrics> =  
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
let tableOverflowPhoto : TableExtractor1<unit> = 
    assertHeaderCell "Overflow Chamber Photo"
    

let extractOutfallInfo : TableExtractor1<OutfallInfo> = 
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
let tableOutfallPhoto : TableExtractor1<unit> = 
    table1 {
        do! assertHeaderCell "Outfall Photo"
        return ()
    }

let scopeOfWorksTable : TableExtractor1<string> = 
    table1 {
        do! assertHeaderCell "Scope of Works"
        let! ans = withCellM (getCellByIndex 3 1) <| getCellText
        return ans
    }


/// General site / survey info     
let sectionSite : DocExtractor<SiteInfo * SurveyInfo> = 
    section 1
        <| docExtract { 
                let! site           = nextTable extractSiteDetails
                let! surveyInfo     = nextTable extractSurveyInfo
                // Doc now has photo tables - don't extract    
                return (site, surveyInfo)
            } 

let sectionOutstation : DocExtractor<OutstationInfo> = 
    section 2 <| nextTable extractOutstationInfo


let ultrasonicTable : DocExtractor<UltrasonicInfo option> = 
    let info = extractUltrasonicInfo1 |>>> Some
    let photos = tableUltrasonicPhotos |>>> (fun () -> None)
    photos <||> info


let sectionUltrasonics : DocExtractor<UltrasonicInfo list>= 
    let stop = lookahead (whiteSpace >>>. pstring "Section 4")
    section 3 <| 
        (manyTill1 ultrasonicTable stop) |>>> (List.choose id)
    

/// list<Overflow Chamber> * list<Chamber Measurements>
/// ignore <Overflow Chamber Photo>

type ChamberTable = 
    | InfoTable of OverflowChamberInfo
    | MetricsTable of OverflowChamberMetrics
    | PhotoTable of unit



let chamberTable : DocExtractor<ChamberTable> = 
    let info    = extractOverflowChamberInfo &|>>> InfoTable
    let metrics = extractOverflowChamberMetrics &|>>> MetricsTable 
    let photo   = tableOverflowPhoto &|>>> PhotoTable
    nextTable (metrics <|||> info <|||> photo) 

let private getOverflowLists (source: ChamberTable list) 
                                : OverflowChamberInfo list * OverflowChamberMetrics list = 
    let rec work ac bc xs = 
        match xs with
        | [] -> (List.rev ac, List.rev bc)
        | InfoTable a :: rest -> work (a::ac) bc rest
        | MetricsTable b :: rest -> work ac (b::bc) rest
        | PhotoTable _ :: rest -> work ac bc rest
    work [] [] source
            
let sectionChambers : DocExtractor<OverflowChamberInfo list * OverflowChamberMetrics list>= 
    let stop = lookahead (whiteSpace >>>. pstring "Section 5")
    section 4 <| 
        (manyTill1 chamberTable stop |>>> getOverflowLists)


let outfallTable : DocExtractor<OutfallInfo option> = 
    let info = extractOutfallInfo &|>>> Some
    let photo = tableOutfallPhoto &|>>> (fun () -> None)
    nextTable (photo <|||> info) <&?> "outfallTable"


/// Note table parser would find finds "OutFall Photos" if we just looked for 
/// "Outfall".
let sectionOutfalls : DocExtractor<OutfallInfo list>= 
    let stop = lookahead (whiteSpace >>>. pstring "Section 6")
    section 5 <| 
        (manyTill outfallTable stop |>>> (List.choose id))
        



/// Single table - Title (1,1) = "Scope of Works", data in cell (r3,c1):
let scopeOfWorks : DocExtractor<string> = 
    section 6 <| nextTable scopeOfWorksTable


/// Single table - Title (1,1) = "Appendix", data in cell (r2,c1):
let appendix : DocExtractor<string> =  
    section 7 
        <| nextTable (withCellM (getCellByIndex 2 1) <| getCellText)


let parseSurvey : DocExtractor<Survey> = 
    docExtract { 
        let! (site, surveyInfo)     = 
            sw "survey-info"      sectionSite      

        let! outstation     = sw "outstation"        sectionOutstation
        let! ultrasonics    = sw "ultrasonics"      sectionUltrasonics
        let! (chambers, metrics)    = 
            sw "chambers"         sectionChambers
        let! outfalls       = sw "outfalls"         sectionOutfalls
        let! scopeText      = sw "scope-of-works"   scopeOfWorks
        let! appendixText   = sw "appendix"         appendix
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
