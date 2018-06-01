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
    pchar ch >>. preturn ()

   
let pstring (str:string) : DocMonad<string> = 
    withInput <| fun s -> 
        try
            let upper = str.Length - 1 
            let ans = s.[0..upper]
            let rest = s.[upper+1..]
            if ans = str then Ok (State rest,ans) else Err "pstring"
        with 
        | ex -> Err "pstring"


let anyChar : DocMonad<char> = 
    satisfy (fun _ -> true) <?> "anyChar"


let newline : DocMonad<char> = 
    let n1 = (fun _ -> '\n')
    pchar '\n' <|> (pstring "\r\n" |>> n1) <|> (pchar '\r' |>> n1)



