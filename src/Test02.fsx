#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
open Microsoft.Office.Interop

#I @"..\packages\FParsec.1.0.3\lib\portable-net45+win8+wp8+wpa81"
#r "FParsec"
#r "FParsecCS"


#load @"DocSoup\Base.fs"
#load @"DocSoup\Monad2.fs"
open DocSoup.Monad2

// Open Fparsec last
open FParsec


let testDoc = @"G:\work\working\Survey1.docx"

let test01 () = 
    let parser1:TextParser<string> = 
        spaces1 >>. many1Till letter spaces1 |>> System.String.Concat
    let proc : DocSoup<_> = fparse parser1 |>>> id
    runOnFileE proc testDoc |> printfn "%s"