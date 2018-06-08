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
        findPatternMany "D??charge" >>>= fun res -> forM res (fun rgn -> focus rgn getText)
    runOnFileE proc testDoc |> Seq.iter (printfn "%A")

let temp01 (ix:int) = 
    let proc : DocSoup<_> = getTableArea (TableAnchor ix)
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
        findCell "SAI Number" true 
    runOnFileE proc testDoc |> printfn "%A"

let test09 () = 
    let proc : DocSoup<_> = 
        findCells "discharge" false 
    runOnFileE proc testDoc |> Seq.iter (printfn "%A")


let test10 () = 
    let proc : DocSoup<_> = 
        findCellsPattern "D??charge"
    runOnFileE proc testDoc |> Seq.iter (printfn "%A")


let test11 () = 
    let proc : DocSoup<_> = 
        findTable "SAI Number" true 
    runOnFileE proc testDoc |> printfn "%A"

let test12 () =
    let proc : DocSoup<_> = 
        findTableAll ["Survey Information"; "Company"; "Field Engineer"] true
    runOnFileE proc testDoc |> printfn "%A"

let test13 () =
    let proc : DocSoup<_> = 
        findTablesAll ["Chamber Measurements"; "Chamber Name"] true
    runOnFileE proc testDoc |> Seq.iter (printfn "%A")


let test14 () =
    let proc : DocSoup<_> = 
        findTablesPatternAll ["C??mber Measurements"; "C??mber Name"]
    runOnFileE proc testDoc |> Seq.iter (printfn "%A")

// In the sample doc the first occurence of "Scope of Works" is not in a table

let test15 () =
    let proc : DocSoup<_> = 
        findTable "Scope of Works" true 
    runOnFileE proc testDoc |> printfn "%A"

let test16 () =
    let proc : DocSoup<_> = 
        findTablesAll ["Scope of Works"] true 
    runOnFileE proc testDoc |> Seq.iter (printfn "%A")

let test17 () =
    let proc : DocSoup<_> = 
        findCell "Scope of Works" true 
    runOnFileE proc testDoc |> printfn "%A"


let test18 () =
    let proc : DocSoup<_> = 
        docSoup { 
            let! sow = findTable "Scope of Works" true 
            let! ans = 
                focusTable sow (
                    findCell "Scope of Works" true >>>= cellBelow >>>= cellBelow >>>= cellText
                )
            return ans
        }
    runOnFileE proc testDoc |> printfn "%A"
