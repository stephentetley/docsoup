// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocSoup.Internal

module Common = 

    open System
    open System.IO
    open System.Text.RegularExpressions


    /// Splits on Environment.NewLine
    let toLines (source:string) : seq<string> = 
        source.Split(separator=[| Environment.NewLine |], options=StringSplitOptions.None) |> Array.toSeq

    /// Joins with Environment.NewLine
    let fromLines (source:seq<string>) : string = 
        String.concat Environment.NewLine source