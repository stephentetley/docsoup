// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


module SurveyRecord

open FSharp.Data

// Favour strings for data (even for dates, etc.).
// Surveys are very much "free text" and there is no guarantee the 
// input data follows any format.

type Survey = 
    { FileName: string 
    
    }
    
[<Literal>]
let SurveySchema = 
    "File Name(string)"


/// Trick - setting Sample to ExportSchema rather than a sample "row" writes the schema as
/// Headers in the output.
type SurveyTable = 
    CsvProvider< Sample = SurveySchema,
                 Schema = SurveySchema,
                 HasHeaders = true >

type SurveyRow = SurveyTable.Row

let csvRow (survey:Survey) : SurveyRow = 
    new SurveyRow( fileName = survey.FileName
                 )


