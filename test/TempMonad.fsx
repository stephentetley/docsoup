// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


#r "netstandard"
#r "System.Xml.Linq"

type ErrMsg = string

type Answer<'a> = Result<'a, ErrMsg>

type RowReader<'a> = RowReader of (string [] -> Answer<'a>)


let inline private apply1 (ma: RowReader<'a>) 
                          (handle: string []): Answer<'a>= 
    let (RowReader f) = ma in f handle

    
let inline mreturn (x:'a) : RowReader<'a> = 
    RowReader <| fun _ -> Ok x

let inline private bindM (ma:RowReader<'a>) 
    (f :'a -> RowReader<'b>) : RowReader<'b> =
        RowReader <| fun handle -> 
            match apply1 ma handle with
            | Error msg -> Error msg
            | Ok a -> apply1 (f a) handle

let inline private zeroM () : RowReader<'a> = 
    RowReader <| fun _ -> Error "zeroM"

/// "First success"
let inline private combineM (ma:RowReader<'a>) 
                            (mb:RowReader<'a>) : RowReader<'a> = 
    RowReader <| fun handle -> 
        match apply1 ma handle with
        | Error msg -> apply1 mb handle
        | Ok a -> Ok a

let inline private delayM (fn:unit -> RowReader<'a>) : RowReader<'a> = 
    bindM (mreturn ()) fn 

type RowReaderBuilder() = 
    member self.Return x = mreturn x
    member self.Bind (p, f) = bindM p f
    member self.Zero () = zeroM ()
    member self.Combine (ma, mb) = combineM ma mb
    member self.Delay fn = delayM fn
    member self.ReturnFrom ma = ma


let (rowReader:RowReaderBuilder) = new RowReaderBuilder()

//let cellMatch (text:string) : RowReader<unit> = 
//    RowReader <| fun handle -> Ok ()

//let cellBody () : RowReader<string> = 
//    RowReader <| fun handle -> 
//        let ans = handle.[state] in Ok (ans, state+1)


let cell (ix:int) : RowReader<string> = 
    RowReader <| fun handle -> 
        let ans = handle.[ix] in Ok ans

let ( |>> ) (ma:RowReader<'a>) (fn:'a -> 'b) : RowReader<'b> = 
    RowReader <| fun handle -> 
        match apply1 ma handle with
        | Error msg -> Error msg
        | Ok ans -> Ok (fn ans)

let matches (s1:string) (s2:string) = s1 = s2

let assertM (ma:RowReader<bool>) : RowReader<unit> = 
    RowReader <| fun handle -> 
        match apply1 ma handle with
        | Error msg -> Error msg
        | Ok ans -> 
            if ans then Ok () else Error "assertM"

/// Design - I don't think anything is gained by making index (cursor) implicit.
/// It already has the disadvantage of impeding random access.
let readSpan : RowReader<string> = 
    rowReader { 
        do! assertM (cell 0 |>> matches "span")
        let! value = cell 1
        return value
    }


