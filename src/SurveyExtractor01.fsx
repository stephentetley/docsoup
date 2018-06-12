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
open DocSoup

// Open Fparsec last
open FParsec
open System.IO

// Favour strings for dat (even for dates, etc.).
// Word surveys are very "free texty".

let getFieldValue (search:string) (matchCase:bool) : DocSoup<string> = 
    let good = findCell search matchCase >>>= cellRight >>>= cellText
    good <||> sreturn ""

let getFieldValuePattern (search:string)  : DocSoup<string> = 
    let good = findCellPattern search >>>= cellRight >>>= cellText
    good <||> sreturn ""


type SurveyHeader = 
    { SiteName: string
      DischargeName: string
      EngineerName: string
      SurveyDate: string }

type OverflowType = SCREENED | UNSCREENED

type ChamberInfo = 
    { OverflowType: OverflowType
      ChamberName: string 
      RoofToInvert: string
      UsFaceToInvert: string
      OverflowToInvert: string
      ScreenToInvert: string
      EmergencyOverflowToInvert: string
      }
    member v.isEmpty = 
        v.ChamberName = "" 
            && v.RoofToInvert = "" 
            && v.UsFaceToInvert = ""
            && v.OverflowToInvert = ""
            && v.ScreenToInvert = ""
            && v.EmergencyOverflowToInvert = ""

type Survey = 
    { SurveyHeader: SurveyHeader
      ChamberInfos: ChamberInfo list }


let extractSiteDetails : DocSoup<string * string> = 
    docSoup { 
        let! t0 = findTable "Site Details" true
        let! ans = 
            focusTable t0 <|  
                tupleM2 (getFieldValue "Site Name" true)
                          (getFieldValue "Discharge Name" true)
        return ans
    }

let extractSurveyInfo : DocSoup<string * string> = 
    docSoup { 
        let! t0 = findTable "Survey Information" true
        let! ans = 
            focusTable t0 <|  
                tupleM2 (getFieldValue "Engineer Name" true)
                          (getFieldValue "Date of Survey" true)
        return ans
    }


let extractSurveyHeader : DocSoup<SurveyHeader> = 
    docSoup { 
        let! (sname,dname) = extractSiteDetails
        let! (engineer, sdate) = extractSurveyInfo
        return { 
            SiteName = sname;
            DischargeName = dname;
            EngineerName = engineer;
            SurveyDate = sdate }
    }

let extractOverflowType (anchor:TableAnchor) : DocSoup<OverflowType> = 
    pipeM2 (DocMonad.optional <| findText "Screen to Invert" false)
            (DocMonad.optional <| findText "Emergency overflow level" false)
            (fun a b -> 
                match a,b with
                | Some _ ,Some _ -> SCREENED
                | _,_ -> UNSCREENED)

let extractChamberInfo1 (anchor:TableAnchor) : DocSoup<ChamberInfo> = 
    let makeInfo (otype:OverflowType) (name:string) (roofDist:string) 
                    (usDist:string) (ovDist:string) (scDist:string) 
                    (emDist:string) : ChamberInfo  = 
        { OverflowType = otype
          ChamberName = name;
          RoofToInvert = roofDist;
          UsFaceToInvert = usDist; 
          OverflowToInvert = ovDist;
          ScreenToInvert = scDist; 
          EmergencyOverflowToInvert = emDist }
    focusTable anchor <| 
        (makeInfo   <&&>  (extractOverflowType anchor)
                    <**> (getFieldValue "Chamber Name" false)
                    <**> (getFieldValue "Roof Slab to Invert" false)
                    <**> (getFieldValue "Transducer Face to Invert" false)
                    <**> (getFieldValue "Overflow level to Invert" false)
                    <**> (getFieldValuePattern "Bottom*Screen*Invert")
                    <**> (getFieldValuePattern "Emergency * Invert") )

let extractChamberInfos : DocSoup<ChamberInfo list>= 
    docSoup { 
        let! anchors = findTables "Chamber Measurement" true
        let! allInfos = mapM extractChamberInfo1 anchors
        return (List.filter (fun (info:ChamberInfo) -> not info.isEmpty) allInfos)
    } <&?> "ChamberInfos"

let parseSurvey : DocSoup<Survey> = 
    pipeM2 extractSurveyHeader 
            extractChamberInfos
            (fun a b -> { SurveyHeader = a; ChamberInfos = b })

let processSurvey (docPath:string) : unit = 
    printfn "Doc: %s" docPath
    runOnFileE parseSurvey docPath |> printfn "%A"


let processSite(folderPath:string) : unit  =
    printfn "Site: '%s'" folderPath
    System.IO.DirectoryInfo(folderPath).GetFiles(searchPattern = "*Survey.docx")
        |> Array.iter (fun (info:System.IO.FileInfo) -> processSurvey info.FullName)

let main () : unit = 
    let root = @"G:\work\Projects\events2\surveys_returned"
    System.IO.DirectoryInfo(root).GetDirectories ()
        |> Array.iter (fun (info:System.IO.DirectoryInfo) -> processSite info.FullName)




