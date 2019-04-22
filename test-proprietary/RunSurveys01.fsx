// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "System.Xml"
#r "System.Xml.Linq"

open System.IO

// Use FSharp.Data for CSV output
#I @"C:\Users\stephen\.nuget\packages\FSharp.Data\3.1.1\lib\netstandard2.0"
#I @"C:\Users\stephen\.nuget\packages\FSharp.Data\3.1.1\typeproviders\fsharp41\netstandard2.0"
#r @"FSharp.Data.dll"
#r @"FSharp.Data.DesignTime"
open FSharp.Data


#load @"SurveyRecord.fs"
#load @"SurveyExtractor.fs"
open SurveyRecord
open SurveyExtractor


//let processSurvey (docPath:string) : unit = 
//    printfn "Doc: %s" docPath
//    runOnFileE parseSurvey docPath |> (fun a -> printfn "%A" (surveyToXml a))


//let processSite(folderPath:string) : unit  =
//    printfn "Site: '%s'" folderPath
//    System.IO.DirectoryInfo(folderPath).GetFiles(searchPattern = "*Survey.docx")
//        |> Array.iter (fun (info:System.IO.FileInfo) -> processSurvey info.FullName)


let main () : unit = 
    let root = @"G:\work\Projects\uqpb\commission-form-samples"
    let files = System.IO.DirectoryInfo(root).GetFiles(searchPattern = "*.docx")
    files |> Seq.iter (printfn "%A")
