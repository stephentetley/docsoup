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
let testpath = @"G:\work\working\Survey1.docx"

let runIt (fn : Word.Document -> 'a) : 'a = 
    let oapp = new Word.ApplicationClass (Visible = true) 
    let odoc = oapp.Documents.Open(FileName = rbox testpath)
    let ans = fn odoc
    odoc.Close(SaveChanges = rbox false)
    oapp.Quit()
    ans


let dummy1 () = 
    let fn (doc : Word.Document) = 
        let dstart = doc.Content.Start
        let dend = doc.Content.End
        doc.Range(rbox dstart, rbox dend).Text
    runIt fn

// Find just finds the sought for text. To be useful for bookmarking
// we would need to see range to the right (and maybe to the the left)
let dummy2 () = 
    let rngfind (doc : Word.Document) = 
        let mutable rng = doc.Range()
        let ans = rng.Find.Execute(FindText = rbox "Contractor Information")
        rng.Text
    runIt rngfind



let dummy3 () = 
    let viewtables (doc : Word.Document) = 
        for table1 in doc.Tables do 
            printfn "(%i,%i)\n>>>>>\n%s<<<<<" table1.Range.Start table1.Range.End table1.Range.Text
    runIt viewtables

let dummy3a () = 
    // Have to cast Tables collection to a Seq...
    let action (doc : Word.Document) = 
        Seq.cast doc.Tables 
        |> Seq.iter (fun (table1 : Word.Table) -> printfn "(%i,%i)\n" table1.Range.Start table1.Range.End)
    runIt action


let dummy4 () = 
    let fn (doc : Word.Document) = 
        let docrng = doc.Range ()
        let mutable rng1 = doc.Range()
        let ans = rng1.Find.Execute(FindText = rbox "Site")
        let rng2 = rightDifference docrng rng1
        let ans2 = match rng2 with
                   | Some r -> sRestOfLine <| r.Text 
                   | None -> "bad range"
        rng1.Text, ans2
    runIt fn




let test01 () = 
    let text = test text testpath
    text

let test02 () = 
    let p1 = extractor { let! a = text
                         return a }
    let text = test p1 testpath
    text

let test03 () = 
    let p1 = withTable 1 <| text
    let text = test p1 testpath
    text

// Hopefully find should be delimited...
// Note though that find isn't very "good" it just finds (or not) what it looks for, we really 
// need a means of using find as a anchor where we can access text (or tables, etc.) next to it.
let test04 () = 
    let p1 = extractor { let! a = restOfLine
                         let! b = find "Contractor Information"
                         let! c = restOfLine
                         return a,b,c }
    test p1 testpath

let anchor (searchtext: string) (p: Extractor<'a>) : Extractor<'a> = failwith "TODO"

let nextTableDown (p: Extractor<'a>) : Extractor<'a> = failwith "TODO"

let freeTextValue (search : string) : Extractor<string> = failwith "TODO"

type GeneralInfo = 
    { Site : string
      LevelControlName : string     // aka "Process Application"
      SiteArea : string }


let extract1 () = 
    let p0 = extractor { let! sn = freeTextValue "Site"
                         let! lcn = freeTextValue "Process Application"
                         let! sa = freeTextValue "Site Area" 
                         return {Site=sn; LevelControlName=lcn; SiteArea=sa} }
    let (p1 : Extractor<GeneralInfo>) = anchor "Survey General Information" <| nextTableDown p0
    test p1 testpath


let anchorU (searchtext: string) : Extractor<'a> = failwith "TODO"


//let nextTableDownU : Extractor<'a> = failwith "TODO"
//
//
//let extract2 () = 
//    let (p1 : Extractor<string>) = anchorU "Survey General Information" *> nextTableDownU *> text
//    test p1 testpath

    