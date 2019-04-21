// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocSoup.Internal

module Old = 

    open FParsec

    exception FatalParseError of string

    type ParsecParser<'ans> = Parser<'ans,unit>


    type FParsecFallback<'a> = 
        | FParsecOk of 'a
        | FallbackText of string

