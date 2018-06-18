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

let run (fn:Word.Document -> 'a) (fileName:string) : 'a = 
    if System.IO.File.Exists (fileName) then
        let app = new Word.ApplicationClass (Visible = false) :> Word.Application
        try 
            let doc = app.Documents.Open(FileName = ref (fileName :> obj))
            let ans = fn doc
            doc.Close(SaveChanges = rbox false)
            app.Quit()
            ans
        with
        | ex -> 
            try 
                app.Quit ()
                failwith ex.Message
            with
            | _ -> failwith ex.Message
                
    else 
        failwith <| sprintf "Cannot find file %s" fileName

let test01 () = 
    let procM (doc:Word.Document) = 
        let tables = doc.Tables |> Seq.cast<Word.Table>
        Seq.iter (fun (table:Word.Table) -> printfn "Table ID: '%s'" table.ID) tables
    run procM testDoc

// is this observably slower than test01?
// NO
let test02 () = 
    let procM (doc:Word.Document) = 
        let tables = doc.Tables 
        let indexes = [1.. tables.Count]
        List.iter (fun (ix:int) -> printfn "Table ID: '%s'" (tables.Item(ix).ID)) indexes
    run procM testDoc


let findPatternMany (doc:Word.Document) (search:string) : Region list =
    let rec work (rng:Word.Range) (ac: Region list) : Region list = 
        rng.Find.Execute (FindText = rbox search, 
                            MatchWildcards = rbox true,
                            MatchCase = rbox true,
                            Forward = rbox true) |> ignore
        if rng.Find.Found then
            let region = extractRegion rng
            work rng (region::ac)
        else
            List.rev ac

    let range1 = doc.Range()
    range1.Find.ClearFormatting ()
    work range1 []

let test03 () = 
    let procM (doc:Word.Document) = 
        let results = findPatternMany doc "O??rflow"
        List.iter (printfn "Region: %A") results
    run procM testDoc

let findPatternManyT1 (doc:Word.Document) (search:string) : Region list =
    let rec work (rng:Word.Range) (ac: Region list) : Region list = 
        rng.Find.Execute (FindText = rbox search, 
                            MatchWildcards = rbox true,
                            MatchCase = rbox true,
                            Forward = rbox true) |> ignore
        if rng.Find.Found then
            let region = extractRegion rng
            work rng (region::ac)
        else
            List.rev ac

    let range1 = doc.Tables.Item(9).Range
    printfn "Initial region: %A" (extractRegion range1)
    range1.Find.ClearFormatting ()
    work range1 []


let test03b () = 
    let procM (doc:Word.Document) = 
        let results = findPatternManyT1 doc "O??rflow"
        List.iter (printfn "Region: %A") results
    run procM testDoc

let test03c () = 
    let procM (doc:Word.Document) = 
        let range1 = doc.Tables.Item(9).Range
        let results = boundedFindPatternMany "Ov??flow" extractRegion range1
        List.iter (printfn "Region: %A") results
    run procM testDoc

let test04 (search:string)  = 
    let procM (doc:Word.Document) = 
        let table1 = doc.Tables.Item(1)
        let range1 = table1.Range
        let rgn1 = extractRegion range1
        let found = range1.Find.Execute (FindText = rbox search, 
                                            MatchWildcards = rbox true,
                                            MatchCase = rbox true,
                                            Forward = rbox true)
        if found then 
            let rgn2 = extractRegion range1
            printfn "r1:%A; r2:%A" rgn1 rgn2 
        else
            printfn "Not found"
    run procM testDoc

let test05 () = 
    let first = TableAnchor.First
    printfn "%A" first.Next.Next


// Does sepBy (optionally) terminate?
let testFParsec01 () = 
    let parser = 
        FParsec.Primitives.parse { 
            let! nums = FParsec.Primitives.sepBy digit (pchar ';') 
            let! letter1 = letter
            return (nums,letter1)
        }
    match FParsec.CharParsers.run parser "1;2;3;4;X" with
    | Success(a, _, _) -> 
        printfn "Success %A" a
    | Failure(_,_,_) -> 
        printfn "sepBy is proper sep (not endBy)"
        FParsec.CharParsers.run parser "1;2;3;4X" |> printfn "Second attempt: %A"

