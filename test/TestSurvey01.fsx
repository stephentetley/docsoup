#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
open Microsoft.Office.Interop

#r "System.Xml"
#r "System.Xml.Linq"

#I @"C:\Users\stephen\.nuget\packages\FParsec\1.0.4-rc3\lib\netstandard1.6"
#r "FParsec"
#r "FParsecCS"

open FParsec
open System.IO


#load @"..\src\DocSoup\Internal\Common.fs"
#load @"..\src\DocSoup\RowExtractor.fs"
#load @"..\src\DocSoup\TablesExtractor.fs"
open DocSoup.RowExtractor
open DocSoup.TablesExtractor

#load @"SurveySyntax.fs"
#load @"SurveyExtractor.fs"
open SurveySyntax
open SurveyExtractor


let processSurvey (docPath:string) : unit = 
    printfn "Doc: %s" docPath
    runOnFileE parseSurvey docPath |> (fun a -> printfn "%A" (surveyToXml a))


let processSite(folderPath:string) : unit  =
    printfn "Site: '%s'" folderPath
    System.IO.DirectoryInfo(folderPath).GetFiles(searchPattern = "*Survey.docx")
        |> Array.iter (fun (info:System.IO.FileInfo) -> processSurvey info.FullName)


let main () : unit = 
    let root = @"G:\work\Projects\events2\surveys_returned"
    System.IO.DirectoryInfo(root).GetDirectories ()
        |> Array.take 2
        |> Array.iteri 
            (fun (ix:int) (info:System.IO.DirectoryInfo) -> processSite info.FullName)




