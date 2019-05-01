// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace Extractors.Usar

[<RequireQualifiedAccess>]
module SurveyV2 =

    open DocSoup
    open Extractors.Usar.Schema

    let extractSurveyInfo : Body.Extractor< {| SiteName: string
                                             ; SensorName: string
                                             ; ProcessArea: string 
                                             ; AssetReference: string |} > = 
        ignoreCase <| Body.findTable (Table.firstCell  &>> Cell.innerTextIsMatch "Site Name") 
            &>> pipeM4 (Table.findNameValue2Row "Site Name")
                       (Table.findNameValue2Row "Sensor name")
                       (Table.findNameValue2Row "Process area")
                       (Table.findNameValue2Row "Asset Reference")
                       (fun siteName sensorName processArea reference -> 
                            {| SiteName = siteName
                             ; SensorName = sensorName
                             ; ProcessArea = processArea 
                             ; AssetReference = reference 
                            |})

    let extractVisitInfo : Body.Extractor< {| Engineer: string
                                             ; SurveyDate: string |} > = 
        ignoreCase <| Body.findTable (Table.firstCell  &>> Cell.innerTextIsMatch "Surveyed By") 
            &>> pipeM2 (Table.findNameValue2Row "Surveyed By")
                       (Table.findNameValue2Row "Date")
                       (fun engineer surveyDate -> 
                            {| Engineer = engineer
                             ; SurveyDate = surveyDate 
                            |})



    let usarSurveyExtractor : Document.Extractor<UsarSurveyRow> = 
        Document.body 
            &>> pipeM2 extractSurveyInfo 
                        extractVisitInfo
                        ( fun r1 r2 -> 
                            UsarSurveyRow   ( siteName = r1.SiteName
                                            , sensorName = r1.SensorName
                                            , processArea = r1.ProcessArea
                                            , assetReference = r1.AssetReference
                                            , engineer = r2.Engineer
                                            , surveyDate = r2.SurveyDate
                                            ))

    let processUsarSurvey (filePath:string) : Answer<UsarSurveyRow>  =
        Document.runExtractor filePath usarSurveyExtractor
