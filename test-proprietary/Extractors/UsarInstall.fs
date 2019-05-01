// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace Extractors

module UsarInstall =

    open System.IO

    open FSharp.Data

    open DocSoup


    let extractGeneralInfo : Body.Extractor< {| SiteName: string
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

    let extractVisitInfo : Body.Extractor< {| Engineer: string
                                            ; InstallDate: string |} > = 
        ignoreCase <| Body.findTable (Table.firstCell  &>> Cell.isMatch "Checked By") 
            &>> pipeM2 (Table.findNameValue2Row "Checked By")
                       (Table.findNameValue2Row "Date")
                       (fun engineer installDate -> 
                            {| Engineer = engineer
                             ; InstallDate = installDate 
                            |})

    [<Literal>]
    let OutputSchema = 
        "Site Name(string), Sensor Name(string), " +
        "Process Area(string), Asset Reference(string), " +
        "Engineer(string), Install Date(string)"

    type UsarInstallTable = 
        CsvProvider< Sample = OutputSchema,
                     Schema = OutputSchema,
                     HasHeaders = true >

    type UsarInstallRow = UsarInstallTable.Row


    let usarInstallExtractor : Document.Extractor<UsarInstallRow> = 
        Document.body 
            &>> pipeM2 extractGeneralInfo 
                        extractVisitInfo
                        ( fun r1 r2 -> 
                            UsarInstallRow  ( siteName = r1.SiteName
                                            , sensorName = r1.SensorName
                                            , processArea = r1.ProcessArea
                                            , assetReference = r1.AssetReference
                                            , engineer = r2.Engineer
                                            , installDate = r2.InstallDate
                                            ))

    let processUsarInstall (filePath:string) : Answer<UsarInstallRow>  =
        Document.runExtractor filePath usarInstallExtractor
