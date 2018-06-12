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



type SurveyHeader = 
    { SiteName: string
      DischargeName: string
      EngineerName: string
      SurveyDate: string }

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

type UltrasonicInfo = 
    { MonitoredDischarge: string 
      ProcessOrFacilityName: string
      Manufacturer: string
      Model: string
      SerialNumber: string }
    member v.isEmpty = 
        v.ProcessOrFacilityName = "" 
            && v.Manufacturer = "" 
            && v.Model = ""
            && v.SerialNumber = ""

type Survey = 
    { SurveyHeader: SurveyHeader
      UltrasonicInfos: UltrasonicInfo list
      ChamberInfos: ChamberInfo list }

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

let extractSiteDetails : DocSoup<string * string> = 
    docSoup { 
        let! t0 = findTable "Site Details" true
        let! ans = 
            focusTable t0 <|  
                tupleM2 (getFieldValue "Site Name" true)
                          (getFieldValue "Discharge Name" true)
        return ans
    }

let extractSurveyInfo : DocSoup<string * string> = 
    docSoup { 
        let! t0 = findTable "Survey Information" true
        let! ans = 
            focusTable t0 <|  
                tupleM2 (getFieldValue "Engineer Name" true)
                          (getFieldValue "Date of Survey" true)
        return ans
    }


let extractSurveyHeader : DocSoup<SurveyHeader> = 
    docSoup { 
        let! (sname,dname) = extractSiteDetails
        let! (engineer, sdate) = extractSurveyInfo
        return { 
            SiteName = sname;
            DischargeName = dname;
            EngineerName = engineer;
            SurveyDate = sdate }
    }


let extractUltrasonicInfo1 (anchor:TableAnchor) : DocSoup<UltrasonicInfo> = 
    let makeInfo (disName:string) (procName:string) (manu:string) 
                    (model:string) (snumber:string) : UltrasonicInfo  = 
        { MonitoredDischarge = disName;
          ProcessOrFacilityName = procName;
          Manufacturer = manu;
          Model = model;
          SerialNumber = snumber }
    focusTable anchor <| 
        (makeInfo   <&&> (getFieldValue "Discharge Being Monitored" false)
                    <**> (getFieldValue "Process or Facility" false)
                    <**> (getFieldValue "Manufacturer" false)
                    <**> (getFieldValue "Model" false)
                    <**> (getFieldValue "Serial Number" false)
                    )

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

let parseSurvey : DocSoup<Survey> = 
    pipeM3 extractSurveyHeader 
            extractUltrasonicInfos
            extractChamberInfos
            (fun a b c -> { SurveyHeader = a; 
                            UltrasonicInfos = b; 
                            ChamberInfos = c})
