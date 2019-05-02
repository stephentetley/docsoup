// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace Extractors.Usar

[<RequireQualifiedAccess>]
module SurveyV1 = 

    open DocSoup
    open Extractors.Usar.Schema

    /// Name value pairs are typically a line of the form 
    /// "<name> : <value>" within the lines found by 'spacedLines'.
    /// In some cases we supply a 'not match' for disambiguation.
    
    let extractNameValue (nameRegex: string) 
                         (notNameRegex: string option) : Text.Extractor<string> = 
        let testName = 
            match notNameRegex with
            | None -> Text.isMatch nameRegex
            | Some notName -> Text.isMatch nameRegex <&&> Text.isNotMatch notName
        let rightOfRegex = sprintf @"%s(\s*:?\s*)" nameRegex
        Text.findLine testName &>> Text.rightOfMatch rightOfRegex


    let extractSurveyInfo : Body.Extractor< {| SiteName: string
                                             ; SensorName: string
                                             ; ProcessArea: string 
                                             ; AssetReference: string |} > = 
        let tableMarkers = [| "Site" ; "Process Application"; "Site Area" |]
        ignoreCase <| Body.findTable (Table.spacedText &>> Text.allMatch tableMarkers) &>> Table.spacedText
            &>> pipeM4 (extractNameValue "Site" (Some "Site Area"))
                       (extractNameValue "Process Application" None)
                       (extractNameValue "Site Area" None)
                       (mreturn "{no asset reference}")
                       (fun siteName sensorName processArea reference -> 
                            {| SiteName = siteName
                             ; SensorName = sensorName
                             ; ProcessArea = processArea 
                             ; AssetReference = reference 
                            |})

    let extractVisitInfo : Body.Extractor< {| Engineer: string
                                            ; SurveyDate: string |} > = 
        let tableMarkers = [| "Surveyed By" ; "Date" |]
        ignoreCase <| Body.findTable (Table.spacedText &>> Text.allMatch tableMarkers) &>> Table.spacedText
            &>> pipeM2 (extractNameValue "Surveyed By" None)
                       (extractNameValue "Date" None)
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

