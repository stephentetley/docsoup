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

/// Utility parsers

let getFieldValue (search:string) (matchCase:bool) : DocSoup<string> = 
    let good = findCell search matchCase >>>= cellRight >>>= cellText
    good <||> sreturn ""

let getFieldValuePattern (search:string)  : DocSoup<string> = 
    let good = findCellPattern search >>>= cellRight >>>= cellText
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

/// Survey parsers

let extractSiteDetails : DocSoup<SiteInfo> = 
    let parser1 = 
        docSoup { 
            let! sname          = getFieldValue "Site Name" false
            let! uid            = getFieldValue "SAI Number2" false
            let! discharge      = getFieldValue "Discharge Name" false
            let! watercourse    = getFieldValue "Receiving Watercourse" false
            return {
                SiteName = sname
                SaiNumber = uid
                DischargeName = discharge
                ReceivingWatercourse = watercourse
                }
        }
    findTable "Site Details" true >>>= fun anchor -> focusTable anchor parser1

let extractSurveyInfo : DocSoup<SurveyInfo> = 
    let parser1 =
        docSoup { 
            let! name       = getFieldValue "Engineer Name" false
            let! sdate      = getFieldValue "Date of Survey" false
            return { 
                EngineerName = name
                SurveyDate = sdate
            }
        }
    findTable "Site Details" true >>>= fun anchor -> focusTable anchor parser1



let extractOutstationInfo : DocSoup<OutstationInfo option> = 
    let parser1 : DocSoup<OutstationInfo> = 
        docSoup { 
            let! name = getFieldValue "Outstation Name" false
            let! rtuAddr = getFieldValue "RTU Address" false
            let! otype = getFieldValue "Outstation Type" false
            let! snumber = getFieldValue "Serial Number" false
            return { 
                OutstationName = name
                RtuAddress = rtuAddr
                OutstationType = otype 
                SerialNumber = snumber } 
        }
    let toOpt (outstation:OutstationInfo) =
        if outstation.isEmpty then None else Some outstation
    findTable "Outstation" false >>>= fun anchor -> 
        focusTable anchor parser1 |>>> toOpt
 
// Focus should already be limited to the table in question.
let extractRelay (relayNumber:int) : DocSoup<RelaySetting> = 
    let funPattern = sprintf "Relay*%i*Function" relayNumber
    let onPattern = sprintf "Relay*%i*On" relayNumber
    let offPattern = sprintf "Relay*%i*Off" relayNumber
    let makeSetting (funId:string) (onSet:string) (offSet:string) : RelaySetting = 
        { RelayNumber = relayNumber
          RelayFunction = funId
          OnSetPoint = onSet
          OffSetPoint = offSet }
    pipeM3 (getFieldValuePattern funPattern)
            (getFieldValuePattern onPattern)
            (getFieldValuePattern offPattern)
            makeSetting

let extractRelays : DocSoup<RelaySetting list> = 
    mapM extractRelay [1..6] |>>> List.filter (fun (r:RelaySetting) -> not r.isEmpty)

let extractUltrasonicInfo1 (anchor:TableAnchor) : DocSoup<UltrasonicInfo> = 
    let makeInfo (disName:string) (procName:string) (manu:string) 
                    (model:string) (snumber:string) (piTag:string)
                    (emptyDist:string) (span:string) 
                    (relays:RelaySetting list) : UltrasonicInfo  = 
        { MonitoredDischarge = disName;
          ProcessOrFacilityName = procName;
          Manufacturer = manu;
          Model = model;
          SerialNumber = snumber;
          PITag = piTag;
          EmptyDistance = emptyDist;
          Span = span;
          Relays = relays}
    focusTable anchor <| 
        (makeInfo   <&&> (getFieldValue "Discharge Being Monitored" false)
                    <**> (getFieldValue "Process or Facility" false)
                    <**> (getFieldValue "Manufacturer" false)
                    <**> (getFieldValue "Model" false)
                    <**> (getFieldValue "Serial Number" false)
                    <**> (getFieldValue "P & I Tag" false)
                    <**> (getFieldValue "Empty Distance" false)
                    <**> (getFieldValue "Span" false)
                    <**> extractRelays )

let extractUltrasonicInfos : DocSoup<UltrasonicInfo list>= 
    let dict: MultiTableParser<UltrasonicInfo> = 
        { GetAnchors = findTables "Ultrasonic Level Control" true;
          TableParser = extractUltrasonicInfo1;
          TestNotEmpty = fun (info:UltrasonicInfo) -> not info.isEmpty }
    parseMultipleTables dict <&?> "UltrasonicInfos"

let extractOverflowType (anchor:TableAnchor) : DocSoup<OverflowType> = 
    pipeM2 (DocMonad.optional <| findText "Screen to Invert" false)
            (DocMonad.optional <| findText "Emergency overflow level" false)
            (fun a b -> 
                match a,b with
                | Some _ ,Some _ -> SCREENED
                | _,_ -> UNSCREENED)

// Note - applicative style parser is not an improvement as we need `makeInfo` 
// which itself is verbose.
let extractChamberInfo1 (anchor:TableAnchor) : DocSoup<ChamberInfo> = 
    let makeInfo (otype:OverflowType) (name:string) (roofDist:string) 
                    (usDist:string) (ovDist:string) (scDist:string) 
                    (emDist:string) : ChamberInfo  = 
        { OverflowType = otype
          ChamberName = name;
          RoofToInvert = roofDist;
          UsFaceToInvert = usDist; 
          OverflowToInvert = ovDist;
          ScreenToInvert = scDist; 
          EmergencyOverflowToInvert = emDist }
    focusTable anchor <| 
        (makeInfo   <&&>  (extractOverflowType anchor)
                    <**> (getFieldValue "Chamber Name" false)
                    <**> (getFieldValue "Roof Slab to Invert" false)
                    <**> (getFieldValue "Transducer Face to Invert" false)
                    <**> (getFieldValue "Overflow level to Invert" false)
                    <**> (getFieldValuePattern "Bottom*Screen*Invert")
                    <**> (getFieldValuePattern "Emergency * Invert") )

let extractChamberInfos : DocSoup<ChamberInfo list>= 
    let dict: MultiTableParser<ChamberInfo> = 
        { GetAnchors = findTables "Chamber Measurement" true;
          TableParser = extractChamberInfo1;
          TestNotEmpty = fun (info:ChamberInfo) -> not info.isEmpty }
    parseMultipleTables dict <&?> "ChamberInfos"

let extractOutfallInfo1 (anchor:TableAnchor) : DocSoup<OutfallInfo> = 
    focusTable anchor <|
        docSoup { 
            let! dname = getFieldValue "Discharge Name" false
            let! gridRef = getFieldValue "Grid Ref" false
            let! proven = getFieldValue "Outfall Proven" false
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
    let parser1 = 
        findCell "Scope of Works" false >>>= cellBelow >>>= cellBelow >>>= cellText
    findTable "Scope of Works"  true >>>= fun anchor -> focusTable anchor parser1

let appendixText : DocSoup<string> = 
    let parser1 = 
        findCell "Appendix" false >>>= cellBelow >>>= cellText
    findTable "Appendix"  true >>>= fun anchor -> focusTable anchor parser1

let parseSurvey : DocSoup<Survey> = 
    docSoup { 
        let! site           = extractSiteDetails
        let! surveyInfo     = extractSurveyInfo
        let! outstation     = extractOutstationInfo
        let! ultrasonics    = extractUltrasonicInfos
        let! chambers       = extractChamberInfos
        let! outfalls       = extractOutfallInfos
        let! scope          = scopeOfWorks <||> sreturn ""
        let! appendix       = appendixText <||> sreturn ""
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
