// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


module SurveyExtractor

open System.IO

open DocSoup

open SurveyRecord

let processSurvey(filePath:string) : Answer<SurveyRow>  =
    let name = FileInfo(filePath).Name
    let tempSurvey = { FileName = name }
    let row = csvRow(tempSurvey)
    Ok row


