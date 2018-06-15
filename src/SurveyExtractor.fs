// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module SurveyExtractor

open DocSoup.Base
open DocSoup.DocMonad
open DocSoup

open SurveySyntax


// Note - layout changes are expected for the input documents we are querying here.
// This makes us favour by-name access rather than (faster) by-index access.



// *************************************
// Helpers

let sw (msg:string) (ma:DocSoup<'focus,'a>) : DocSoup<'focus,'a> =
    docSoup { 
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
        focusCellM (findCell search matchCase >>>= cellRight) <| getCellText
    good <||> sreturn ""

/// Returns "" if no cell matches the search.
/// This speeds things up a bit, but for our use case here we are
/// concerned about layout changes.
let getFieldValueByRow (row:int) : TableExtractor<string> = 
    let good = 
        focusCellM (getCellByIndex { RowIx = row; ColumnIx = 2 }) <| getCellText
    good <||> sreturn ""


/// Returns "" if no cell matches the search.
let getFieldValuePattern (search:string) : TableExtractor<string> = 
    let good = 
        focusCellM (findCellPattern search >>>= cellRight) <| getCellText
    good <||> sreturn ""

type MultiTableParser<'answer> = 
    { GetAnchors: DocExtractor<TableAnchor list>
      TableParser: TableExtractor<'answer>
      TestNotEmpty: 'answer -> bool }

let parseMultipleTables (dict:MultiTableParser<'answer>) : DocExtractor<'answer list>= 
    docSoup { 
        let! anchors = dict.GetAnchors
        let! allAnswers = mapM (fun anchor -> focusTable anchor dict.TableParser) anchors
        return (List.filter dict.TestNotEmpty allAnswers)
    } <&?> "parseMultipleTables"



// *************************************
// Survey parsers

let extractSiteDetails : DocExtractor<SiteInfo> = 
    focusTableM (getTableByIndex 1) <| 
        docSoup { 
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


let extractSurveyInfo : DocExtractor<SurveyInfo> = 
    focusTableM (getTableByIndex 2) <| 
        docSoup { 
            let! name       = getFieldValue "Engineer Name" false
            let! sdate      = getFieldValue "Date of Survey" false
            return { 
                EngineerName = name
                SurveyDate = sdate
            }
        }


let extractOutstationInfo : DocExtractor<OutstationInfo option> = 
    let toOpt (outstation:OutstationInfo) =
        if outstation.isEmpty then None else Some outstation

    let parser1 : TableExtractor<OutstationInfo> = 
        docSoup { 
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
    focusTableM (findTable "Outstation" true) (parser1 |>>> toOpt)
 
// Focus should already be limited to the table in question.
let extractRelay (relayNumber:int) : TableExtractor<RelaySetting> = 
    let funPattern  = sprintf "Relay*%i*Function" relayNumber
    let onPattern   = sprintf "Relay*%i*On" relayNumber
    let offPattern  = sprintf "Relay*%i*Off" relayNumber
    docSoup { 
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
    mapM extractRelay [1..6] |>>> List.filter (fun (r:RelaySetting) -> not r.isEmpty)

let extractUltrasonicMonitorInfo1 : TableExtractor<UltrasonicMonitorInfo> = 
    docSoup {
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
    docSoup {
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

let extractUltrasonicInfo1 (levelTable:TableAnchor) : DocExtractor<UltrasonicInfo> = 
    docSoup {
        let! monitor = 
            focusTable levelTable extractUltrasonicMonitorInfo1
        let! sensor = 
            focusTableM (nextTable levelTable) extractUltrasonicSensorInfo1
        return { 
            MonitorInfo = monitor
            SensorInfo = sensor } 
    }


let extractUltrasonicInfos : DocExtractor<UltrasonicInfo list>= 
    let dict: MultiTableParser<UltrasonicInfo> = 
        { GetAnchors = findTables "Ultrasonic Level Control" true;
          TableParser = extractUltrasonicMonitorInfo1;
          TestNotEmpty = fun (info:UltrasonicInfo) -> not info.isEmpty }
    parseMultipleTables dict <&?> "UltrasonicInfos"

let extractOverflowType : TableExtractor<OverflowType> =  
    docSoup { 
        let! a = DocMonad.optional <| findText "Screen to Invert" false
        let! b = DocMonad.optional <| findText "Emergency overflow level" false
        match a,b with
        | Some _, Some _ -> return SCREENED
        | _, _ -> return UNSCREENED
    }

// Note - applicative style parser is not an improvement as we need `makeInfo` 
// which itself is verbose.
let extractChamberInfo1 : TableExtractor<ChamberInfo> =  
    docSoup { 
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


let extractChamberInfos : DocExtractor<ChamberInfo list>= 
    let dict: MultiTableParser<ChamberInfo> = 
        { GetAnchors = findTables "Chamber Measurement" true;
          TableParser = extractChamberInfo1;
          TestNotEmpty = fun (info:ChamberInfo) -> not info.isEmpty }
    parseMultipleTables dict <&?> "ChamberInfos"

let extractOutfallInfo1 : TableExtractor<OutfallInfo> = 
    docSoup { 
        let! dname      = getFieldValue "Discharge Name" false
        let! gridRef    = getFieldValue "Grid Ref" false
        let! proven     = getFieldValue "Outfall Proven" false
        return { 
            DischargeName = dname
            OutfallGridRef = gridRef
            OutfallProven = proven
        }
    }


// Note table parser would find finds "OutFall Photos" if we just looked for 
// "Outfall".
let extractOutfallInfos : DocExtractor<OutfallInfo list>= 
    let dict: MultiTableParser<OutfallInfo> = 
        { GetAnchors = findTables "Outfall Proven" false
          TableParser = extractOutfallInfo1
          TestNotEmpty = fun (info:OutfallInfo) -> not info.isEmpty }
    parseMultipleTables dict <&?> "OutfallInfos"

/// Single table - Title (1,1) = "Scope of Works", data in cell (r3,c1):
let scopeOfWorks : DocExtractor<string> = 
    focusTableM (findTable "Scope of Works" true) <<
        focusCellM (getCellByIndex { RowIx = 3; ColumnIx = 1 }) <| getCellText

/// Single table - Title (1,1) = "Appendix", data in cell (r2,c1):
let appendixText : DocExtractor<string> =         
    focusTableM (findTable "Appendix" true) <<
        focusCellM (getCellByIndex { RowIx = 2; ColumnIx = 1 }) <| getCellText
     


let parseSurvey : DocExtractor<Survey> = 
    docSoup { 
        let! site           = sw "site"             extractSiteDetails
        let! surveyInfo     = sw "surveyInfo"       extractSurveyInfo
        let! outstation     = sw "outstation"       extractOutstationInfo
        let! ultrasonics    = sw "ultrasonics"      extractUltrasonicInfos
        let! chambers       = sw "chambers"         extractChamberInfos
        let! outfalls       = sw "outfalls"         extractOutfallInfos
        let! scope          = sw "scope"            (scopeOfWorks <||> sreturn "")
        let! appendix       = sw "appendix"         (appendixText <||> sreturn "")
        return { 
            SiteDetails = site
            SurveyInfo = surveyInfo
            OutstationInfo = outstation
            UltrasonicInfos = ultrasonics
            ChamberInfos = chambers
            OutfallInfos = outfalls
            ScopeOfWorks = scope
            AppendixText = appendix
            }
    }
