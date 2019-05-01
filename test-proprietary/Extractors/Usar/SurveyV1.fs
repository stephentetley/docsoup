// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace Extractors.Usar

[<RequireQualifiedAccess>]
module SurveyV1 = 

    open DocSoup
    open Extractors.Usar.Schema


    let extractSurveyInfo : Body.Extractor< {| SiteName: string
                                             ; SensorName: string
                                             ; ProcessArea: string 
                                             ; AssetReference: string |} > = 
        let tableMarkers = [| "Site" ; "Process Application"; "Site Area" |]
        ignoreCase <| Body.findTable (Table.innerTextAllMatch tableMarkers ) 
            &>> pipeM4 (Table.innerText)
                       (mreturn "TODO")
                       (mreturn "TODO")
                       (mreturn "{no asset reference}")
                       (fun siteName sensorName processArea reference -> 
                            {| SiteName = siteName
                             ; SensorName = sensorName
                             ; ProcessArea = processArea 
                             ; AssetReference = reference 
                            |})
