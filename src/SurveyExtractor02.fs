// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module SurveyExtractor02

open DocSoup.Base
open DocSoup.TableExtractor
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
let getFieldValue (search:string) (matchCase:bool) : TableExtractor<string> = 
    let good = 
        focusCellM (findCell search matchCase &>>= cellRight) <| getCellText
    good <|||> treturn ""

/// Returns "" if no cell matches the search.
/// This speeds things up a bit, but for our use case here we are
/// concerned about layout changes.
let getFieldValueByRow (row:int) : TableExtractor<string> = 
    let good = 
        focusCellM (getCellByIndex row 2) <| getCellText
    good <|||> treturn ""


/// Returns "" if no cell matches the search.
let getFieldValuePattern (search:string) : TableExtractor<string> = 
    let good = 
        focusCellM (findCellByPattern search &>>= cellRight) <| getCellText
    good <|||> treturn ""

//type MultiTableParser<'answer> = 
//    { GetAnchors: DocExtractor<TableAnchor list>
//      TableParser: TableExtractor<'answer>
//      TestNotEmpty: 'answer -> bool }

//let parseMultipleTables (dict:MultiTableParser<'answer>) : DocExtractor<'answer list>= 
//    docSoup { 
//        let! anchors = dict.GetAnchors
//        let! allAnswers = mapM (fun anchor -> focusTable anchor dict.TableParser) anchors
//        return (List.filter dict.TestNotEmpty allAnswers)
//    } <&?> "parseMultipleTables"



// *************************************
// Survey parsers

let extractSiteDetails : TableExtractor<SiteInfo> = 
    tableExtract { 
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


let extractSurveyInfo : TableExtractor<SurveyInfo> = 
    tableExtract { 
        let! name       = getFieldValue "Engineer Name" false
        let! sdate      = getFieldValue "Date of Survey" false
        return { 
            EngineerName = name
            SurveyDate = sdate
        }
    }


let extractOutstationInfo : TableExtractor<OutstationInfo> = 
    tableExtract { 
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
let extractRelay (relayNumber:int) : TableExtractor<RelaySetting> = 
    let funPattern  = sprintf "Relay*%i*Function" relayNumber
    let onPattern   = sprintf "Relay*%i*On" relayNumber
    let offPattern  = sprintf "Relay*%i*Off" relayNumber
    tableExtract { 
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


let extractRelays : TableExtractor<RelaySetting list> = 
    DocSoup.TableExtractor.mapM extractRelay [1..6] ||>>> List.filter (fun (r:RelaySetting) -> not r.isEmpty)

let extractUltrasonicMonitorInfo1 : TableExtractor<UltrasonicMonitorInfo> = 
    tableExtract {
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

let extractUltrasonicSensorInfo1 : TableExtractor<UltrasonicSensorInfo> = 
    tableExtract {
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

//let extractUltrasonicInfo1 (levelTable:TableAnchor) : DocExtractor<UltrasonicInfo> = 
//    docSoup {
//        let! monitor = 
//            focusTable levelTable extractUltrasonicMonitorInfo1
//        let! sensor = 
//            focusTableM (nextTable levelTable) extractUltrasonicSensorInfo1
//        return { 
//            MonitorInfo = monitor
//            SensorInfo = sensor } 
//    }


//let extractUltrasonicInfos : DocExtractor<UltrasonicInfo list>= 
//    let dict: MultiTableParser<UltrasonicInfo> = 
//        { GetAnchors = findTables "Ultrasonic Level Control" true;
//          TableParser = extractUltrasonicMonitorInfo1;
//          TestNotEmpty = fun (info:UltrasonicInfo) -> not info.isEmpty }
//    parseMultipleTables dict <&?> "UltrasonicInfos"

let extractOverflowType : TableExtractor<OverflowType> =  
    tableExtract { 
        let! a = TableExtractor.optional <| findCell "Screen to Invert" false
        let! b = TableExtractor.optional <| findCell "Emergency overflow level" false
        match a,b with
        | Some _, Some _ -> return SCREENED
        | _, _ -> return UNSCREENED
    }


let extractChamberInfo1 : TableExtractor<ChamberInfo> =  
    tableExtract { 
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


//let extractChamberInfos : DocExtractor<ChamberInfo list>= 
//    let dict: MultiTableParser<ChamberInfo> = 
//        { GetAnchors = findTables "Chamber Measurement" true;
//          TableParser = extractChamberInfo1;
//          TestNotEmpty = fun (info:ChamberInfo) -> not info.isEmpty }
//    parseMultipleTables dict <&?> "ChamberInfos"

let extractOutfallInfo1 : TableExtractor<OutfallInfo> = 
    tableExtract { 
        let! dname      = getFieldValue "Discharge Name" false
        let! gridRef    = getFieldValue "Grid Ref" false
        let! proven     = getFieldValue "Outfall Proven" false
        return { 
            DischargeName = dname
            OutfallGridRef = gridRef
            OutfallProven = proven
        }
    }


//// Note table parser would find finds "OutFall Photos" if we just looked for 
//// "Outfall".
//let extractOutfallInfos : DocExtractor<OutfallInfo list>= 
//    let dict: MultiTableParser<OutfallInfo> = 
//        { GetAnchors = findTables "Outfall Proven" false
//          TableParser = extractOutfallInfo1
//          TestNotEmpty = fun (info:OutfallInfo) -> not info.isEmpty }
//    parseMultipleTables dict <&?> "OutfallInfos"

/// Single table - Title (1,1) = "Scope of Works", data in cell (r3,c1):
let scopeOfWorks : DocExtractor<string> = 
    docExtract { 
        do! advanceM (findText "Section 6" false)
        let! text           = 
            sw "scope-of-works"       <| nextTable (focusCellM (getCellByIndex 3 1) <| getCellText)
        return text
    }


/// Single table - Title (1,1) = "Appendix", data in cell (r2,c1):
let appendix : DocExtractor<string> =  
    docExtract { 
        do! advanceM (findText "Section 7" false)
        let! text           = 
            sw "appendix"       <| nextTable (focusCellM (getCellByIndex 2 1) <| getCellText)
        return text
    }

/// General site / survey info     
let section1 : DocExtractor<SiteInfo * SurveyInfo> = 
    docExtract { 
        let! site           = 
            sw "site"       <| nextTable extractSiteDetails
        let! surveyInfo     = 
            sw "surveyInfo" <| nextTable extractSurveyInfo
        // Doc now has photo tables - don't extract    
        return (site, surveyInfo)
    } 

let section2 : DocExtractor<OutstationInfo> = 
    docExtract { 
        do! advanceM (findText "Section 2" false)
        let! outstation     = 
            sw "outstation"         <| nextTable extractOutstationInfo
        return outstation
    }


let parseSurvey : DocExtractor<Survey> = 
    docExtract { 
        let! (site, surveyInfo)     = section1      

        let! outstation             = section2

        let! ultrasonics    = sw "ultrasonics"      extractUltrasonicInfos
        let! chambers       = sw "chambers"         extractChamberInfos
        let! outfalls       = sw "outfalls"         extractOutfallInfos
        let! scopeText      = scopeOfWorks
        let! appendixText   = appendix
        return { 
            SiteDetails = site
            SurveyInfo = surveyInfo
            OutstationInfo = outstation
            UltrasonicInfos = ultrasonics
            ChamberInfos = chambers
            OutfallInfos = outfalls
            ScopeOfWorks = scopeText
            AppendixText = appendixText
            }
    }
