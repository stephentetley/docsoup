// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace Extractors.Usar

[<AutoOpen>]
module Schema = 

    open FSharp.Data

    [<Literal>]
    let SurveySchema = 
        "Site Name(string), Sensor Name(string), " +
        "Process Area(string), Asset Reference(string), " +
        "Engineer(string), Survey Date(string)"

    type UsarSurveyTable = 
        CsvProvider< Sample = SurveySchema,
                     Schema = SurveySchema,
                     HasHeaders = true >

    type UsarSurveyRow = UsarSurveyTable.Row


    [<Literal>]
    let InstallSchema = 
        "Site Name(string), Sensor Name(string), " +
        "Process Area(string), Asset Reference(string), " +
        "Engineer(string), Install Date(string)"

    type UsarInstallTable = 
        CsvProvider< Sample = InstallSchema,
                     Schema = InstallSchema,
                     HasHeaders = true >

    type UsarInstallRow = UsarInstallTable.Row


