#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
open Microsoft.Office.Interop

#I @"..\packages\FParsec.1.0.3\lib\portable-net45+win8+wp8+wpa81"
#r "FParsec"
#r "FParsecCS"


#load @"DocSoup\Base.fs"
#load @"DocSoup\Monad2.fs"
open DocSoup.Base
open DocSoup.Monad2

// Open Fparsec last
open FParsec


let testDoc = @"G:\work\working\Survey1.docx"

let test01 () = 
    let parser1:TextParser<string> = 
        spaces1 >>. many1Till letter spaces1 |>> System.String.Concat
    let proc : DocSoup<_> = fparse parser1 |>>> id
    runOnFileE proc testDoc |> printfn "%s"

let test02 () = 
    let proc : DocSoup<_> = findText "Discharge" true
    runOnFileE proc testDoc |> printfn "%A"

let test03 () = 
    let proc : DocSoup<_> = findTextMany "Discharge" true
    runOnFileE proc testDoc |> List.iter (printfn "%A")

let test04 () = 
    let proc : DocSoup<_> = tableAreas
    runOnFileE proc testDoc |> Seq.iter (printfn "%A")

let test05 () = 
    let proc : DocSoup<_> = findPattern "D??charge"
    runOnFileE proc testDoc |> printfn "%A"

let test06 () = 
    let proc : DocSoup<_> = 
        findPatternMany "D??charge" >>>= fun res -> mapM res (fun rgn -> focus rgn getText)
    runOnFileE proc testDoc |> Seq.iter (printfn "%A")

let temp01 (ix:int) = 
    let proc : DocSoup<_> = getTableArea (TableId ix)
    runOnFileE proc testDoc |> printfn "%A"


let test07 () = 
    let proc : DocSoup<_> = 
        findText "SAI Number" true >>>= containingTable >>>=  getTableArea >>>= fun rgn -> focus rgn getText
    runOnFileE proc testDoc |> printfn "%s"


// there is an occurence of "Site Details" prior to the "Site Details" table.
let test07a () = 
    let proc : DocSoup<_> = 
        findText "Site Details" true
    runOnFileE proc testDoc |> printfn "%A"


let test07b () = 
    let proc : DocSoup<_> = 
        findText "SAI Number" true
    runOnFileE proc testDoc |> printfn "%A"

let test08 () = 
    let proc : DocSoup<_> = 
        findText "SAI Number" true >>>= containingCell
    runOnFileE proc testDoc |> printfn "%A"