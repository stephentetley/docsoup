module DocSoup.CharParsers

open DocSoup.DocMonad

// We need the regular parser combinators for data extraction, e.g. int, 
// float, digit, letter, etc.


let satisfy (test:char -> bool) : DocMonad<char> = 
    withInput <| fun s -> 
        try
            let a = s.[0]
            let rest = s.[1..]
            if test a then Ok (State rest,a) else Err "satisfy"
        with 
        | ex -> Err "satisfy"


let pchar (ch:char) : DocMonad<char> = 
    satisfy (fun c1 -> c1 = ch) <?> "pchar"


let skipChar (ch:char) : DocMonad<unit> = 
    swapError "skipChar" <| 
        pchar ch >>. preturn ()




let anyChar : DocMonad<char> = 
    satisfy (fun _ -> true) <?> "anyChar"

let anyOf (chars:seq<char>) : DocMonad<char> = 
    swapError "anyOf" <| 
        satisfy (fun ch -> Seq.exists (fun c1 -> c1 = ch) chars)


let noneOf (chars:seq<char>) : DocMonad<char> = 
    swapError "noneOf" <| 
        satisfy (fun ch -> not <| Seq.exists (fun c1 -> c1 = ch) chars)

let asciiLower:  DocMonad<char> = anyOf ['a'..'z']
let asciiUpper:  DocMonad<char> = anyOf ['A'..'Z']
let asciiLetter: DocMonad<char> = asciiLower <|> asciiUpper

let lower:  DocMonad<char> = satisfy System.Char.IsLower
let upper:  DocMonad<char> = satisfy System.Char.IsUpper
let letter: DocMonad<char> = satisfy System.Char.IsLetter

let digit: DocMonad<char> = satisfy System.Char.IsDigit




// Parsing strings
// ===============

let pstring (str:string) : DocMonad<string> = 
    withInput <| fun s -> 
        try
            let upper = str.Length - 1 
            let ans = s.[0..upper]
            let rest = s.[upper+1..]
            if ans = str then Ok (State rest,ans) else Err "pstring"
        with 
        | ex -> Err "pstring"

// compares with CurrentCultureIgnoreCase
let pstringCI (str:string) : DocMonad<string> = 
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


let manySatisfy (test:char -> bool) : DocMonad<string> = 
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

let manySatisfy2 (test1:char -> bool) (test2:char -> bool) : DocMonad<string> = 
    let good = 
        docMonad { 
            let! a = satisfy test1
            let! b = manySatisfy test2
            return (string a + b)
        }
    let bad = preturn ""
    good <|> bad



let many1Satisfy (test:char -> bool) : DocMonad<string> = 
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


let many1Satisfy2 (test1:char -> bool) (test2:char -> bool) : DocMonad<string> = 
    docMonad { 
        let! a = satisfy test1
        let! b = many1Satisfy test2
        return (string a + b)
    }



// Parsing whitespace
// ==================

let tab : DocMonad<char> = pchar '\t'


let newline : DocMonad<char> = 
    let rOptn = pchar '\r' .>> optionalz (pchar '\n')
    (pchar '\n' <|> rOptn) |>> fun _ -> '\n'

let spaces : DocMonad<unit> = 
    manySatisfy System.Char.IsWhiteSpace |>> ignore

let spaces1 : DocMonad<unit> = 
    many1Satisfy System.Char.IsWhiteSpace |>> ignore

let eof : DocMonad<unit> = 
    withInput <| fun s -> 
        if s = "" then Ok (State s, ()) else Err "eof (not empty)"

// TODO - whats the preferred way of char[] -> string?
let anyString (ntimes:int32) : DocMonad<string> = 
    count ntimes (newline <|> anyChar) |>> fun arr -> System.String.Concat(arr)