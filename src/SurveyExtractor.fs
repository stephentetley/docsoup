// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module SurveyExtractor

open DocSoup.Base
open DocSoup.DocMonad
open DocSoup


open FParsec
open System.IO


// Favour strings for data (even for dates, etc.).
// Word surveys are very "free texty" and there is no guarantee the input data 
// follows any format.

// *************************************
// Syntax tree

type SiteInfo = 
    { SiteName: string
      SaiNumber: string
      DischargeName: string 
      ReceivingWatercourse: string }

type SurveyInfo = 
    { EngineerName: string
      SurveyDate: string }

type OutstationInfo = 
    { OutstationName: string
      RtuAddress: string
      OutstationType: string
      SerialNumber: string }
    member v.isEmpty = 
        v.OutstationName = ""
            && v.RtuAddress = ""
            && v.OutstationType = ""
            && v.SerialNumber = ""


// Store relay number inside the data. This means we can have a sparse list of
// relays.
type RelaySetting = 
    { RelayNumber: int
      RelayFunction: string
      OnSetPoint: string
      OffSetPoint: string }
    member v.isEmpty = 
        v.RelayFunction = "" 
            && v.OnSetPoint = ""
            && v.OffSetPoint = ""


type UltrasonicInfo = 
    { MonitoredDischarge: string 
      ProcessOrFacilityName: string
      Manufacturer: string
      Model: string
      SerialNumber: string
      PITag: string
      EmptyDistance: string
      Span: string
      Relays: RelaySetting list }
    member v.isEmpty = 
        v.ProcessOrFacilityName = "" 
            && v.Manufacturer = "" 
            && v.Model = ""
            && v.SerialNumber = ""
            && v.PITag = ""
            && v.EmptyDistance = ""
            && v.Span = ""
            && v.Relays.IsEmpty


type OverflowType = SCREENED | UNSCREENED

type ChamberInfo = 
    { OverflowType: OverflowType
      ChamberName: string 
      RoofToInvert: string
      UsFaceToInvert: string
      OverflowToInvert: string
      ScreenToInvert: string
      EmergencyOverflowToInvert: string
      }
    member v.isEmpty = 
        v.ChamberName = "" 
            && v.RoofToInvert = "" 
            && v.UsFaceToInvert = ""
            && v.OverflowToInvert = ""
            && v.ScreenToInvert = ""
            && v.EmergencyOverflowToInvert = ""

type OutfallInfo = 
    { DischargeName: string 
      OutfallGridRef: string
      OutfallProven: string }
    member v.isEmpty = 
        v.OutfallGridRef = ""
            && v.OutfallProven = ""

type Survey = 
    { SiteDetails: SiteInfo
      SurveyInfo: SurveyInfo
      OutstationInfo: option<OutstationInfo>
      UltrasonicInfos: UltrasonicInfo list
      ChamberInfos: ChamberInfo list 
      OutfallInfos: OutfallInfo list
      ScopeOfWorks: string 
      AppendixText: string }

// *************************************
// Helpers

let sw (msg:string) (ma:DocSoup<'a>) : DocSoup<'a> =
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
let getFieldValue (search:string) (matchCase:bool) : DocSoup<string> = 
    let good = findCell search matchCase >>>= cellRight >>>= getCellText
    good <||> sreturn ""

/// Returns "" if no cell matches the search.
let getFieldValuePattern (search:string) : DocSoup<string> = 
    let good = findCellPattern search >>>= cellRight >>>= getCellText
    good <||> sreturn ""

type MultiTableParser<'answer> = 
    { GetAnchors: DocSoup<TableAnchor list>
      TableParser: TableAnchor -> DocSoup<'answer>
      TestNotEmpty: 'answer -> bool }

let parseMultipleTables (dict:MultiTableParser<'answer>) : DocSoup<'answer list>= 
    docSoup { 
        let! anchors = dict.GetAnchors
        let! allAnswers = mapM dict.TableParser anchors
        return (List.filter dict.TestNotEmpty allAnswers)
    } <&?> "parseMultipleTables"



// *************************************
// Survey parsers

let extractSiteDetails : DocSoup<SiteInfo> = 
    focusTableM (findTable "Site Details" true) <| 
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


let extractSurveyInfo : DocSoup<SurveyInfo> = 
    focusTableM (findTable "Site Details" true) <| 
        docSoup { 
            let! name       = getFieldValue "Engineer Name" false
            let! sdate      = getFieldValue "Date of Survey" false
            return { 
                EngineerName = name
                SurveyDate = sdate
            }
        }


let extractOutstationInfo : DocSoup<OutstationInfo option> = 
    let toOpt (outstation:OutstationInfo) =
        if outstation.isEmpty then None else Some outstation

    let parser1 : DocSoup<OutstationInfo> = 
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
let extractRelay (relayNumber:int) : DocSoup<RelaySetting> = 
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


let extractRelays : DocSoup<RelaySetting list> = 
    mapM extractRelay [1..6] |>>> List.filter (fun (r:RelaySetting) -> not r.isEmpty)

let extractUltrasonicInfo1 (anchor:TableAnchor) : DocSoup<UltrasonicInfo> = 
    focusTable anchor <| 
        docSoup {
            let! disName = getFieldValue "Discharge Being Monitored" false
            let! procName = getFieldValue "Process or Facility" false
            let! manuf = getFieldValue "Manufacturer" false
            let! model = getFieldValue "Model" false
            let! snumber = getFieldValue "Serial Number" false
            let! piTag = getFieldValue "P & I Tag" false
            let! emptyDist = getFieldValue "Empty Distance" false
            let! span = getFieldValue "Span" false
            let! relays =  extractRelays
            return { 
                MonitoredDischarge = disName
                ProcessOrFacilityName = procName
                Manufacturer = manuf
                Model = model
                SerialNumber = snumber
                PITag = piTag
                EmptyDistance = emptyDist
                Span = span
                Relays = relays
            }
          } 


let extractUltrasonicInfos : DocSoup<UltrasonicInfo list>= 
    let dict: MultiTableParser<UltrasonicInfo> = 
        { GetAnchors = findTables "Ultrasonic Level Control" true;
          TableParser = extractUltrasonicInfo1;
          TestNotEmpty = fun (info:UltrasonicInfo) -> not info.isEmpty }
    parseMultipleTables dict <&?> "UltrasonicInfos"

let extractOverflowType (anchor:TableAnchor) : DocSoup<OverflowType> = 
    focusTable anchor <| 
        docSoup { 
            let! a = DocMonad.optional <| findText "Screen to Invert" false
            let! b = DocMonad.optional <| findText "Emergency overflow level" false
            match a,b with
            | Some _, Some _ -> return SCREENED
            | _, _ -> return UNSCREENED
        }

// Note - applicative style parser is not an improvement as we need `makeInfo` 
// which itself is verbose.
let extractChamberInfo1 (anchor:TableAnchor) : DocSoup<ChamberInfo> = 
    focusTable anchor <| 
        docSoup { 
            let! otype      = extractOverflowType anchor
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


let extractChamberInfos : DocSoup<ChamberInfo list>= 
    let dict: MultiTableParser<ChamberInfo> = 
        { GetAnchors = findTables "Chamber Measurement" true;
          TableParser = extractChamberInfo1;
          TestNotEmpty = fun (info:ChamberInfo) -> not info.isEmpty }
    parseMultipleTables dict <&?> "ChamberInfos"

let extractOutfallInfo1 (anchor:TableAnchor) : DocSoup<OutfallInfo> = 
    focusTable anchor <|
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
let extractOutfallInfos : DocSoup<OutfallInfo list>= 
    let dict: MultiTableParser<OutfallInfo> = 
        { GetAnchors = findTables "Outfall Proven" false
          TableParser = extractOutfallInfo1
          TestNotEmpty = fun (info:OutfallInfo) -> not info.isEmpty }
    parseMultipleTables dict <&?> "OutfallInfos"

let scopeOfWorks : DocSoup<string> = 
    focusTableM (findTable "Scope of Works" true) <| 
        findCell "Scope of Works" false >>>= cellBelow >>>= cellBelow >>>= getCellText

let appendixText : DocSoup<string> =         
    focusTableM (findTable "Appendix" true) <|
        findCell "Appendix" false >>>= cellBelow >>>= getCellText


let parseSurvey : DocSoup<Survey> = 
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
