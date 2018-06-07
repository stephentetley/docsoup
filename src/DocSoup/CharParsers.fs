// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocSoup.CharParsers

open DocSoup.DocMonad

// We need the regular parser combinators for data extraction, e.g. int, 
// float, digit, letter, etc.
// It actually might be nicer to "embed" FParsec as it is already optimized.
// This is to consider later.


let satisfy (test:char -> bool) : DocParser<char> = 
    withInput <| fun s -> 
        try
            let a = s.[0]
            let rest = s.[1..]
            if test a then Ok (State rest,a) else Err "satisfy"
        with 
        | ex -> Err "satisfy"


let pchar (ch:char) : DocParser<char> = 
    satisfy (fun c1 -> c1 = ch) <?> "pchar"


let skipChar (ch:char) : DocParser<unit> = 
    swapError "skipChar" <| 
        pchar ch >>. preturn ()




let anyChar : DocParser<char> = 
    satisfy (fun _ -> true) <?> "anyChar"

let anyOf (chars:seq<char>) : DocParser<char> = 
    swapError "anyOf" <| 
        satisfy (fun ch -> Seq.exists (fun c1 -> c1 = ch) chars)


let noneOf (chars:seq<char>) : DocParser<char> = 
    swapError "noneOf" <| 
        satisfy (fun ch -> not <| Seq.exists (fun c1 -> c1 = ch) chars)

let asciiLower:  DocParser<char> = anyOf ['a'..'z']
let asciiUpper:  DocParser<char> = anyOf ['A'..'Z']
let asciiLetter: DocParser<char> = asciiLower <|> asciiUpper

let lower:  DocParser<char> = satisfy System.Char.IsLower
let upper:  DocParser<char> = satisfy System.Char.IsUpper
let letter: DocParser<char> = satisfy System.Char.IsLetter

let digit: DocParser<char> = satisfy System.Char.IsDigit




// Parsing strings
// ===============

let pstring (str:string) : DocParser<string> = 
    withInput <| fun s -> 
        try
            let upper = str.Length - 1 
            let ans = s.[0..upper]
            let rest = s.[upper+1..]
            if ans = str then Ok (State rest,ans) else Err "pstring"
        with 
        | ex -> Err "pstring"

// compares with CurrentCultureIgnoreCase
let pstringCI (str:string) : DocParser<string> = 
    withInput <| fun s -> 
        try
            let upper = str.Length - 1 
            let ans = s.[0..upper]
            let rest = s.[upper+1..]
            if System.String.Equals(ans,str, System.StringComparison.CurrentCultureIgnoreCase) then 
                Ok (State rest,ans) 
            else Err "pstringCI"
        with 
        | ex -> Err "pstringCI"


let manySatisfy (test:char -> bool) : DocParser<string> = 
    withInput <| fun s -> 
        try
            let mutable ix = 0
            while test s.[ix] do
                ix <- ix + 1
                
            let ans,rest = 
                if ix > 0 then 
                    s.[0..ix-1], s.[ix..]
                else 
                    "", s
            Ok (State rest,ans)
        with 
        | ex -> Err "manySatisfy"

let manySatisfy2 (test1:char -> bool) (test2:char -> bool) : DocParser<string> = 
    let good = 
        docParse { 
            let! a = satisfy test1
            let! b = manySatisfy test2
            return (string a + b)
        }
    let bad = preturn ""
    good <|> bad



let many1Satisfy (test:char -> bool) : DocParser<string> = 
    withInput <| fun s -> 
        try
            let mutable ix = 0
            while test s.[ix] do
                ix <- ix + 1
                
            if ix > 0 then 
                Ok (State s.[ix..], s.[0..ix-1])
            else 
                Err "many1Satisfy"
        with 
        | ex -> Err "many1Satisfy"


let many1Satisfy2 (test1:char -> bool) (test2:char -> bool) : DocParser<string> = 
    docParse { 
        let! a = satisfy test1
        let! b = many1Satisfy test2
        return (string a + b)
    }



// Parsing whitespace
// ==================

let tab : DocParser<char> = pchar '\t'

let crlf : DocParser<unit> = 
    pchar '\r' >>. pchar '\n' >>. preturn ()

let newline : DocParser<char> = 
    let rOptn = pchar '\r' .>> optionalz (pchar '\n')
    (pchar '\n' <|> rOptn) |>> fun _ -> '\n'

let newlinez : DocParser<unit> = newline |>> ignore


let spaces : DocParser<unit> = 
    let white1 = satisfy System.Char.IsWhiteSpace |>> ignore
    many (crlf <|> white1) |>> ignore

let spaces1 : DocParser<unit> = 
    let white1 = satisfy System.Char.IsWhiteSpace |>> ignore
    many1 (crlf <|> white1) |>> ignore

let eof : DocParser<unit> = 
    withInput <| fun s -> 
        if s = "" then Ok (State s, ()) else Err "eof (not empty)"

let anyString (ntimes:int32) : DocParser<string> = 
    count ntimes (newline <|> anyChar) |>> fun arr -> System.String.Concat(arr)

let restOfLine (skipNewLine:bool) : DocParser<string> = 
    manyTill anyChar (eof <|> newlinez) >>= fun ans ->
    let ans1 = System.String.Concat(ans)
    if skipNewLine then
        optionalz newline >>= fun _ -> preturn ans1
    else preturn ans1
    
        