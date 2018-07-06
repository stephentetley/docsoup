// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module SurveySyntax


open System.Xml
open System.Xml.Linq
open System


// Favour strings for data (even for dates, etc.).
// Our word surveys are very "free texty" and there is no guarantee the 
// input data follows any format.


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


type UltrasonicMonitorInfo = 
    { MonitoredDischarge: string 
      ProcessOrFacilityName: string
      MonitorManufacturer: string
      MonitorModel: string
      SerialNumber: string
      PITag: string
      EmptyDistance: string
      Span: string
      Relays: RelaySetting list }
    member v.isEmpty = 
        v.ProcessOrFacilityName = "" 
            && v.MonitorManufacturer = "" 
            && v.MonitorModel = ""
            && v.SerialNumber = ""
            && v.PITag = ""
            && v.EmptyDistance = ""
            && v.Span = ""
            && v.Relays.IsEmpty

            
type UltrasonicSensorInfo = 
    { SensorManufacturer: string
      SensorModel: string
      SerialNumber: string
      LocationOfSensor: string
      GridRef: string }
    member v.isEmpty = 
        v.SensorManufacturer = "" 
            && v.SensorModel = "" 
            && v.SerialNumber = ""
            && v.LocationOfSensor = ""
            && v.GridRef = ""

type UltrasonicInfo = 
    { MonitorInfo: UltrasonicMonitorInfo
      SensorInfo:  UltrasonicSensorInfo  } 
    member v.isEmpty = 
        v.MonitorInfo.isEmpty 
            && v.SensorInfo.isEmpty

type OverflowChamberInfo = 
    { DischargeName: string
      ChamberName: string
      OverflowGridRef: string
      IsScreened: string } 
    member v.isEmpty  = 
        v.ChamberName = ""
            && v.OverflowGridRef = ""
            && v.IsScreened = ""


type OverflowType = SCREENED | UNSCREENED

type OverflowChamberMetrics = 
    { OverflowType: OverflowType
      ChamberName: string 
      RoofToInvert: string
      UsFaceToInvert: string
      CoverLevelToInvert: string
      OverflowToInvert: string
      ScreenToInvert: string
      EmergencyOverflowToInvert: string
      }
    member v.isEmpty = 
        v.ChamberName = "" 
            && v.RoofToInvert = "" 
            && v.UsFaceToInvert = ""
            && v.CoverLevelToInvert = ""
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
      OutstationInfo: OutstationInfo
      UltrasonicInfos: UltrasonicInfo list
      ChamberInfos: OverflowChamberInfo list
      ChamberMetrics: OverflowChamberMetrics list
      OutfallInfos: OutfallInfo list
      ScopeOfWorks: string 
      AppendixText: string }



let private xname (s:string) : XName = XName.Get s
let private xelement (s:string) :XElement = XElement(xname s)
let private xattribute (s:string) (v:obj) :XAttribute = XAttribute(xname s, v)


let private siteInfoToXml (info:SiteInfo) : XElement = 
    let elt = XElement(xname "SiteInfo")
    elt.Add([   XElement(xname "SiteName", info.SiteName);
                XElement(xname "SaiNumber", info.SaiNumber);
                XElement(xname "DischargeName", info.DischargeName);
                XElement(xname "ReceivingWatercourse", info.ReceivingWatercourse) ])
    elt

let private surveyInfoToXml (info:SurveyInfo) : XElement = 
    let elt = XElement(xname "SurveyInfo")
    elt.Add([   XElement(xname "EngineerName", info.EngineerName);
                XElement(xname "SurveyDate", info.SurveyDate) ])
    elt

let private outstationInfoToXml (info:OutstationInfo) : XElement = 
    let elt = XElement(xname "OutstationInfo")
    elt.Add([   XElement(xname "OutstationName", info.OutstationName);
                XElement(xname "RtuAddress", info.RtuAddress);
                XElement(xname "OutstationType", info.OutstationType);
                XElement(xname "SerialNumber", info.SerialNumber) ])
    elt


/// RelaySetting
let private relaySettingInfoToXml (info:RelaySetting) : XElement = 
    let elt = XElement(xname "RelaySetting")
    elt.Add (xattribute "Number" info.RelayNumber)
    elt.Add([   XElement(xname "RelayFunction", info.RelayFunction);
                XElement(xname "OnSetPoint", info.OnSetPoint);
                XElement(xname "OffSetPoint", info.OffSetPoint) ])
    elt

let private ultrasonicMonitorInfoToXml (info:UltrasonicMonitorInfo) : XElement = 
    let elt = XElement(xname "UltrasonicMonitor")
    elt.Add([   XElement(xname "ProcessOrFacilityName", info.ProcessOrFacilityName);
                XElement(xname "Manufacturer", info.MonitorManufacturer);
                XElement(xname "Model", info.MonitorModel);
                XElement(xname "SerialNumber", info.SerialNumber) ])
    elt.Add(List.map relaySettingInfoToXml info.Relays)
    elt

let private ultrasonicSensorInfoToXml (info:UltrasonicSensorInfo) : XElement = 
    let elt = XElement(xname "UltrasonicSensor")
    elt.Add([   XElement(xname "Manufacturer", info.SensorManufacturer)
                XElement(xname "Model", info.SensorModel);
                XElement(xname "SerialNumber", info.SerialNumber);
                XElement(xname "LocationOfSensor", info.LocationOfSensor);
                XElement(xname "GridRef", info.GridRef)  ])
    elt
    
let private ultrasonicInfoToXml (info:UltrasonicInfo) : XElement = 
    let elt = XElement(xname "UltrasonicInfo")
    elt.Add([   ultrasonicMonitorInfoToXml info.MonitorInfo;
                ultrasonicSensorInfoToXml info.SensorInfo ])
    elt    

let private overflowChamberInfoToXml (info:OverflowChamberInfo) : XElement = 
    let elt = XElement(xname "OverflowChamber")
    elt.Add([   XElement(xname "DischargeName", info.DischargeName);
                XElement(xname "ChamberName", info.ChamberName);
                XElement(xname "OverflowGridRef", info.OverflowGridRef);
                XElement(xname "IsScreened", info.IsScreened)  ])
    elt    

let private overflowChamberMetricsToXml (info:OverflowChamberMetrics) : XElement = 
    let elt = XElement(xname "OverflowMetrics")
    elt.Add([   XElement(xname "ChamberName", info.ChamberName);
                XElement(xname "OverflowType", info.OverflowType.ToString());
                XElement(xname "RoofToInvert", info.RoofToInvert);
                XElement(xname "UsFaceToInvert", info.UsFaceToInvert);
                XElement(xname "CoverLevelToInvert", info.CoverLevelToInvert); 
                XElement(xname "ScreenToInvert", info.ScreenToInvert);
                XElement(xname "OverflowToInvert", info.OverflowToInvert);
                XElement(xname "EmergencyOverflowToInvert", info.EmergencyOverflowToInvert) ])
    elt    

let private outfallInfoToXml (info:OutfallInfo) : XElement = 
    let elt = XElement(xname "Outfall")
    elt.Add([   XElement(xname "DischargeName", info.DischargeName);
                XElement(xname "OutfallGridRef", info.OutfallGridRef);
                XElement(xname "OutfallProven", info.OutfallProven) ])
    elt    

let private scopeOfWorksToXml (body:string) : XElement = 
    XElement(xname "ScopeOfWorks", body)
   
let private appendixToXml (body:string) : XElement = 
    XElement(xname "Appendix", body)



let surveyToXml (survey:Survey) : XDocument  = 
    let doc = XDocument ()
    let root = xelement "EdmSurvey"
    root.Add (xattribute "ProcessingTimeStamp" (DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")))
    doc.Add(root)
    
    root.Add([  siteInfoToXml survey.SiteDetails; 
                surveyInfoToXml survey.SurveyInfo;
                outstationInfoToXml survey.OutstationInfo ])
    root.Add(List.map ultrasonicInfoToXml survey.UltrasonicInfos)
    root.Add(List.map overflowChamberInfoToXml survey.ChamberInfos)
    root.Add(List.map overflowChamberMetricsToXml survey.ChamberMetrics)
    root.Add(List.map outfallInfoToXml survey.OutfallInfos)
    root.Add([  scopeOfWorksToXml survey.ScopeOfWorks; 
                appendixToXml survey.AppendixText ])
    doc



