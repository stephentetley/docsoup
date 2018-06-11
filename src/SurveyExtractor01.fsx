#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
open Microsoft.Office.Interop

#I @"..\packages\FParsec.1.0.3\lib\portable-net45+win8+wp8+wpa81"
#r "FParsec"
#r "FParsecCS"


#load @"DocSoup\Base.fs"
#load @"DocSoup\DocMonad.fs"
open DocSoup.Base
open DocSoup.DocMonad

// Open Fparsec last
open FParsec
open System.IO



let processSurvey (docPath:string) : unit = 
    let proc : DocSoup<_> = 
        findTable "Site Details" true 
    printfn "Doc: %s" docPath
    runOnFileE proc docPath |> printfn "%A"


let processSite(folderPath:string) : unit  =
    printfn "Site: '%s'" folderPath
    System.IO.DirectoryInfo(folderPath).GetFiles(searchPattern = "*Survey.docx")
        |> Array.iter (fun (info:System.IO.FileInfo) -> processSurvey info.FullName)

let main () : unit = 
    let root = @"G:\work\Projects\events2\surveys_returned"
    System.IO.DirectoryInfo(root).GetDirectories ()
        |> Array.iter (fun (info:System.IO.DirectoryInfo) -> processSite info.FullName)




