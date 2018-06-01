#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
open Microsoft.Office.Interop

open System.IO


#load @"DocSoup\Base.fs"
#load @"Old\RangeOperations.fs"
#load @"Old\Extractors.fs"
open DocSoup.Base
open Old.RangeOperations
open Old.Extractors

// Note to self - this example is not "properly structured" tables are free text
//
let testpath = @"G:\work\working\test-range.docx"

let runIt (fn : Word.Document -> 'a) : 'a = 
    let oapp = new Word.ApplicationClass (Visible = true) 
    let odoc = oapp.Documents.Open(FileName = rbox testpath)
    let ans = fn odoc
    odoc.Close(SaveChanges = rbox false)
    oapp.Quit()
    ans

// Word Ranges start at 0


let dummy1 () = 
    let fn (doc : Word.Document) = 
        let dstart = doc.Content.Start
        let dend = doc.Content.End
        let s = doc.Range(rbox dstart, rbox dend).Text
        printfn "%i,%i: \"%s\"" dstart dend s
    runIt fn

let dummy2 () = 
    let fn (doc : Word.Document) = 
        let dstart = 1
        let dend = 2
        let s = doc.Range(rbox dstart, rbox dend).Text
        printfn "%i,%i: \"%s\"" dstart dend s
    runIt fn



let dummy3 () = 
    let fn (doc : Word.Document) = 
        let dstart = 0
        let dend = 1
        let s = doc.Range(rbox dstart, rbox dend).Text
        printfn "%i,%i: \"%s\"" dstart dend s
    runIt fn


let extractRange (a:int) (b:int) (doc:Word.Document) = 
    let rng1 : Word.Range = doc.Range(rbox a, rbox b)
    let rng2 = rng1.Duplicate
    rng2

let dummy4 () =  
    let fn (doc : Word.Document) = 
        let rng1 = extractRange 0 4 doc
        let rng2 = extractRange 3 7 doc
        let oans = rightDifference rng1 rng2 
        match oans with
        | Some rng -> printfn "%i,%i: \"%s\"" rng.Start rng.End rng.Text
        | None -> printfn "None"
    runIt fn