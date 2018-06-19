// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module SurveySyntax



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
      OutstationInfo: OutstationInfo
      UltrasonicInfos: UltrasonicInfo list
      ChamberInfos: OverflowChamberInfo list
      ChamberMetrics: OverflowChamberMetrics list
      OutfallInfos: OutfallInfo list
      ScopeOfWorks: string 
      AppendixText: string }

