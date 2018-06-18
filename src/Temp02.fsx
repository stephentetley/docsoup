#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"
open Microsoft.Office.Interop

#I @"..\packages\FParsec.1.0.3\lib\portable-net45+win8+wp8+wpa81"
#r "FParsec"
#r "FParsecCS"
open FParsec

#load @"DocSoup\Base.fs"
#load @"DocSoup\TableExtractor.fs"
#load @"DocSoup\DocExtractor.fs"
open DocSoup.Base
open DocSoup.TableExtractor
open DocSoup.DocExtractor

let testDoc = @"G:\work\working\Survey1.docx"



let test01 () = 
    let procM : DocExtractor<_> = 
        whiteSpace >>>. parseString "Event"
        
    runOnFileE procM testDoc |> printfn "%A"
