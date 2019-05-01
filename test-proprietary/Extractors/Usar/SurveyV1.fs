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
        ignoreCase <| Body.findTable (Table.firstCell  &>> Cell.isMatch "Site Name") 
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
