#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
open Microsoft.Office.Interop

#load @"DocSoup\Base.fs"
#load @"DocSoup\DocMonad.fs"
#load @"DocSoup\CharParsers.fs"
open DocSoup.Base
open DocSoup.DocMonad
open DocSoup.CharParsers

// Note to self, this test doc is not "well formed". 
// Textual table data is often not split into rows and columns.
let testDoc = @"G:\work\working\Survey1.docx"


let showTable (t1 : Word.Table) = 
    printfn "Rows %i, Columns %i" t1.Rows.Count t1.Columns.Count
    let ans = t1.ConvertToText (rbox Word.WdSeparatorType.wdSeparatorHyphen)
    printfn "Table: %s" ans.Text

let shorten (s:string) = if s.Length > 10 then s.[0..9]+"..." else s

let test01 () = 
    let proc : DocMonad<unit> = 
        docMonad { 
            do! fmapM (printfn "Sections: %i")      countSections
            do! fmapM (printfn "Paragraphs: %i")    countParagraphs
            do! fmapM (printfn "Tables: %i")        countTables
        }
    runOnFileE proc testDoc


let test02 () = 
    let proc (doc:Word.Document) : unit = 
        doc.Tables 
            |> Seq.cast<Word.Table> 
            |> Seq.iter showTable
        doc.Sections 
            |> Seq.cast<Word.Section> 
            |> Seq.iter (fun s1 -> printfn "Tables: %i" s1.Range.Tables.Count)

        let all : Word.Range = doc.Content
        all.Select()
        printfn "Characters: %i" all.Characters.Count
    runOnFileE (liftGlobalOperation proc) testDoc

let test03 () = 
    let proc (doc:Word.Document) : unit = 
        let t1 = doc.Tables.[1]
        printfn "Rows %i, Columns %i" t1.Rows.Count t1.Columns.Count
        printfn "%s" (t1.ConvertToText(rbox Word.WdSeparatorType.wdSeparatorHyphen).Text)
    runOnFileE (liftGlobalOperation proc) testDoc

let test04 () = 
    let proc (doc:Word.Document) : unit = 
        let t1 = doc.Tables.[1]
        let mutable rng1 = t1.Range
        // The mutated range matches what is found.
        let found = rng1.Find.Execute(FindText = rbox "Process Application")
        if found then
            printfn "'%s'" rng1.Text
        else printfn "no found"
    runOnFileE (liftGlobalOperation proc) testDoc





let test05 () = 
    let proc = tupleM2 countTables countSections
    printfn "%A" <| runOnFileE proc testDoc

// All text of the document
let test06 () = 
    printfn "%A" <| runOnFileE cleanText testDoc

// This is nice and high level...
let test07 () = 
    let proc = 
        sequenceM   [ table 1 <| cell (0,0) cleanText
                    ; table 3 <| cell (0,0) cleanText
                    ; table 3 <| cell (1,0) cleanText
                    ; table 3 <| cell (2,0) cleanText
                    ; table 3 <| cell (3,0) cleanText
                    ; table 3 <| cell (4,0) cleanText
                    ; table 3 <| cell (5,0) cleanText
                    ; table 3 << cell (6,0) <| cleanText
                    ]
   
    printfn "%A" <| runOnFileE proc testDoc


let test08 () = 
    let proc = docMonad { 
        let! i = countTables
        let! xs = mapTablesWith (fmapM shorten cleanText)
        return (i,xs)
    }
    printfn "%A" <| runOnFileE proc testDoc

let test09 () = 
    let proc = docMonad { 
        let! (i,xs) = table 3 <| tupleM2 (countCells) (mapCellsWith (fmapM shorten cleanText))
        return (i,xs)
    }
    printfn "%A" <| runOnFileE proc testDoc

let test10 () = 
    let proc : DocMonad<_> = newline >>. pstring "Event"
    runOnFileE proc testDoc |> printfn "%A"
