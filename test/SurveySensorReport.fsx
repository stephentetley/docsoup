// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

#r "netstandard"
#r "System.Xml.Linq"

#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
open Microsoft.Office.Interop

#I @"C:\Users\stephen\.nuget\packages\FParsec\1.0.4-rc3\lib\netstandard1.6"
#r "FParsec"
#r "FParsecCS"

open System

#load @"..\src\DocSoup\Internal\Common.fs"
#load @"..\src\DocSoup\RowExtractor.fs"
#load @"..\src\DocSoup\TablesExtractor.fs"
#load @"SurveySyntax.fs"
#load @"SurveyExtractor.fs"
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

let headings : string list = 
    [ "Site Name"
    ; "Sai Number"
    ; "Discharge Name"
    ; "US Manufacturer"
    ; "US Model"
    ; "Sensor Head Location"
    ; "Sensor Head Grid Ref" ]

let titles : string = String.concat "," headings



let optQuote(s:string) : string = 
    if String.length s > 0 && s.[0] <> '"' then 
        if s.Contains "," then
            sprintf "\"%s\"" s
        else
            s
    else
        s


let csvRow (sensor1:SensorRecord) : string = 
    let cols = 
        [ sensor1.SiteName
        ; sensor1.SaiNumber
        ; sensor1.DischargeName
        ; sensor1.UsManufacturer
        ; sensor1.UsManufacturer
        ; optQuote (sensor1.SensorHeadLocation)
        ; sensor1.SensorHeadGridRef
        ]
    String.concat "," cols



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


let processSurvey (sw:IO.StreamWriter) (docPath:string) : unit = 
    printfn "Doc: %s" docPath
    try 
        runOnFileE parseSurvey docPath 
            |> extractSensorRecords
            |> List.iter (fun r -> sw.WriteLine(csvRow r))
    with
    | _ -> 
        printfn "Fail: %s" docPath


/// All surveys in one folder
let processSurveys (sw:IO.StreamWriter) (folderPath:string) : unit  =
    System.IO.DirectoryInfo(folderPath).GetFiles(searchPattern = "*Survey.docx")
        |> Array.iter (fun (info:System.IO.FileInfo) -> 
                        processSurvey sw info.FullName)

let main () : unit = 
    let root = @"G:\work\Projects\events2\data\surveys_to_read"
    let outPath = @"G:\work\Projects\events2\report_sensor_locations.csv"
    use sw = new IO.StreamWriter(outPath)
    sw.WriteLine titles
    processSurveys sw root




