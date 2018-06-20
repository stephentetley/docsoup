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
#load @"DocSoup\RowExtractor.fs"
#load @"DocSoup\TableExtractor1.fs"
#load @"DocSoup\TablesExtractor.fs"
open DocSoup
open DocSoup.RowExtractor
open DocSoup.TablesExtractor

let testDoc = @"G:\work\working\Survey1.docx"
let testDoc2 = @"G:\work\working\Survey2.docx"


let test01 () = 
    let procM : TablesExtractor<_> = 
        manyTill (parseTable TableExtractor1.getTableRegion) eof
    runOnFileE procM testDoc |> printfn "Ans: '%A'"

let table1 : RowExtractor<string * string> = 
    parseRows { 
        let! title  = row (cellAnything)
        let! name   = row (skip &>>>. cellAnything) 
        return title, name
    }


let test02 () = 
    let procM : TablesExtractor<_> = 
        parseTableRowwise table1
    runOnFileE procM testDoc |> printfn "Ans: '%A'"

