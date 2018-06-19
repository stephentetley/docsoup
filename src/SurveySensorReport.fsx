﻿// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
open Microsoft.Office.Interop

#I @"..\packages\FParsec.1.0.3\lib\portable-net45+win8+wp8+wpa81"
#r "FParsec"
#r "FParsecCS"

#load @"DocSoup\Base.fs"
#load @"DocSoup\TableExtractor1.fs"
#load @"DocSoup\TablesExtractor.fs"
#load @"SurveySyntax.fs"
#load @"SurveyExtractor.fs"
open DocSoup.TableExtractor1
open DocSoup.TablesExtractor
open SurveySyntax
open SurveyExtractor

type SensorRecord = 
    { SiteName: string
      SaiNumber: string
      DischargeName: string
      UsManufacturer: string
      UsModel: string
      SensorHeadLocation: string
      SensorHeadGridRef: string }

let extractSensorRecords (survey:Survey) : SensorRecord list = 
    let siteName = survey.SiteDetails.SiteName
    let saiNumber = survey.SiteDetails.SaiNumber
    let makeRecord (usInfo:UltrasonicInfo) : SensorRecord = 
        { SiteName = siteName
          SaiNumber = saiNumber
          DischargeName = usInfo.MonitorInfo.MonitoredDischarge
          UsManufacturer = usInfo.MonitorInfo.MonitorManufacturer
          UsModel = usInfo.MonitorInfo.MonitorModel
          SensorHeadLocation = usInfo.SensorInfo.LocationOfSensor
          SensorHeadGridRef = usInfo.SensorInfo.GridRef }
    List.map makeRecord survey.UltrasonicInfos


let processSurvey (docPath:string) : unit = 
    printfn "Doc: %s" docPath
    try 
        runOnFileE parseSurvey docPath 
            |> extractSensorRecords
            |> List.iter (printfn "%A")
    with
    | _ -> 
        printfn "Fail: %s" docPath


let processSite(folderPath:string) : unit  =
    printfn "Site: '%s'" folderPath
    System.IO.DirectoryInfo(folderPath).GetFiles(searchPattern = "*Survey.docx")
        |> Array.iter (fun (info:System.IO.FileInfo) -> processSurvey info.FullName)

let main () : unit = 
    let root = @"G:\work\Projects\events2\surveys_returned"
    System.IO.DirectoryInfo(root).GetDirectories ()
        |> Array.iteri 
            (fun (ix:int) (info:System.IO.DirectoryInfo) -> processSite info.FullName)




