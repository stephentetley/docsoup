// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module Combinators = 


    open DocSoup

    /// Operator for fmap.
    let ( |>> ) (ma:ExtractMonad<'a, 'handle>) (fn:'a -> 'b) : ExtractMonad<'b, 'handle> = 
        fmapM fn ma

    /// Flipped fmap.
    let ( <<| ) (fn:'a -> 'b) (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<'b, 'handle> = 
        fmapM fn ma

    /// Perform an action, but ignore its answer.
    let ignoreM (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<unit, 'handle> = 
        ma |>> fun _ -> ()


    let ( <?> ) (ma:ExtractMonad<'a, 'handle>) (msg:string) : ExtractMonad<'a, 'handle> = 
        swapError msg ma


    /// Haskell Applicative's (<*>)
    let apM (mf:ExtractMonad<'a ->'b, 'handle>) (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<'b, 'handle> = 
        mf >>= fun fn -> 
        ma >>= fun a  -> 
        mreturn (fn a)

    /// Perform two actions in sequence. 
    /// Ignore the results of the second action if both succeed.
    let seqL (ma:ExtractMonad<'a, 'handle>) (mb:ExtractMonad<'b, 'handle>) : ExtractMonad<'a, 'handle> = 
        ma >>= fun ans -> 
        mb >>= fun _ -> 
        mreturn ans

    /// Perform two actions in sequence. 
    /// Ignore the results of the first action if both succeed.
    let seqR (ma:ExtractMonad<'a, 'handle>) (mb:ExtractMonad<'b, 'handle>) : ExtractMonad<'b, 'handle> = 
        ma >>= fun _ -> 
        mb >>= fun ans -> 
        mreturn ans

    /// Operator for seqL
    let (.>>) (ma:ExtractMonad<'a, 'handle>) 
              (mb:ExtractMonad<'b, 'handle>) : ExtractMonad<'a, 'handle> = 
        seqL ma mb

    /// Operator for seqR
    let (>>.) (ma:ExtractMonad<'a, 'handle>) 
              (mb:ExtractMonad<'b, 'handle>) : ExtractMonad<'b, 'handle> = 
        seqR ma mb


    let kleisliL (mf:'a -> ExtractMonad<'b, 'handle>)
                 (mg:'b -> ExtractMonad<'c, 'handle>)
                 (source:'a) : ExtractMonad<'c, 'handle> = 
        mf source >>= fun a1 ->
        mg a1 >>= fun ans ->
        mreturn ans


    /// Flipped kleisliL
    let kleisliR (mf:'b -> ExtractMonad<'c, 'handle>)
                 (mg:'a -> ExtractMonad<'b, 'handle>)
                 (source:'a) : ExtractMonad<'c, 'handle> = 
        kleisliL mg mf source



    /// Operator for kleisliL
    let (>=>) (mf : 'a -> ExtractMonad<'b, 'handle>)
              (mg : 'b -> ExtractMonad<'c, 'handle>)
              (source:'a) : ExtractMonad<'c, 'handle> = 
        kleisliL mf mg source


    /// Operator for kleisliR
    let (<=<) (mf : 'b -> ExtractMonad<'c, 'handle>)
              (mg : 'a -> ExtractMonad<'b, 'handle>)
              (source:'a) : ExtractMonad<'c, 'handle> = 
        kleisliR mf mg source


    let andM (ma:ExtractMonad<bool, 'handle>)
             (mb:ExtractMonad<bool, 'handle>) : ExtractMonad<bool, 'handle> =
        ma >>= fun a ->  
        mb >>= fun b -> 
        mreturn (a && b)


    let ( <&&> ) (ma:ExtractMonad<bool, 'handle>)
                 (mb:ExtractMonad<bool, 'handle>) : ExtractMonad<bool, 'handle> =
        andM ma mb


    let orM (ma:ExtractMonad<bool, 'handle>)
            (mb:ExtractMonad<bool, 'handle>) : ExtractMonad<bool, 'handle> =
        ma >>= fun ans -> 
        if ans then 
            mreturn true 
        else 
            mb >>= fun ans2 -> mreturn ans2

    let ( <||> ) (ma:ExtractMonad<bool, 'handle>)
                 (mb:ExtractMonad<bool, 'handle>) : ExtractMonad<bool, 'handle> =
        orM ma mb

    let liftAssert (failMsg:string) (condition:bool) : ExtractMonad<unit, 'handle> = 
        if condition then mreturn () else extractError failMsg

    /// Lift an option value.
    /// Some ans becomes a success value in the ExtractMonad.
    /// None throws a monadic error (not system system exception). 
    let liftOption (opt:'a option) : ExtractMonad<'a, 'handle> = 
        match opt with
        | Some a -> mreturn a
        | None -> extractError "liftOption - None"


    /// Lift an operation that may fail (e.g. an 'IO' operation).
    /// If the action does fail, replace the hard error with 
    /// a (soft) error within the monad.
    let liftOperation (errMsg:string) (operation: unit -> 'a) : ExtractMonad<'a, 'handle> = 
        try 
            let ans = operation ()
            mreturn ans
        with
        | _ -> extractError errMsg


    let assertM (failMsg:string) (cond:ExtractMonad<bool, 'handle>) : ExtractMonad<unit, 'handle> = 
        cond >>= fun ans -> 
        match ans with
        | true -> mreturn ()
        | false -> extractError failMsg


    // liftM (which is fmap)
    let liftM (fn:'a -> 'x) (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<'x, 'handle> = 
        fmapM fn ma

    let liftM2 (fn:'a -> 'b -> 'x) 
               (ma:ExtractMonad<'a, 'handle>) 
               (mb:ExtractMonad<'b, 'handle>) : ExtractMonad<'x, 'handle> = 
        ma >>= fun a ->
        mb >>= fun b ->
        mreturn (fn a b)

    let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
               (ma:ExtractMonad<'a, 'handle>) 
               (mb:ExtractMonad<'b, 'handle>) 
               (mc:ExtractMonad<'c, 'handle>) : ExtractMonad<'x, 'handle> = 
        ma >>= fun a ->
        mb >>= fun b ->
        mc >>= fun c ->
        mreturn (fn a b c)

    let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
               (ma:ExtractMonad<'a, 'handle>) 
               (mb:ExtractMonad<'b, 'handle>) 
               (mc:ExtractMonad<'c, 'handle>) 
               (md:ExtractMonad<'d, 'handle>) : ExtractMonad<'x, 'handle> = 
        ma >>= fun a ->
        mb >>= fun b ->
        mc >>= fun c ->
        md >>= fun d ->
        mreturn (fn a b c d)


    let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) 
               (ma:ExtractMonad<'a, 'handle>) 
               (mb:ExtractMonad<'b, 'handle>) 
               (mc:ExtractMonad<'c, 'handle>) 
               (md:ExtractMonad<'d, 'handle>) 
               (me:ExtractMonad<'e, 'handle>) : ExtractMonad<'x, 'handle> = 
        ma >>= fun a ->
        mb >>= fun b ->
        mc >>= fun c ->
        md >>= fun d ->
        me >>= fun e ->
        mreturn (fn a b c d e)
        

    let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
               (ma:ExtractMonad<'a, 'handle>) 
               (mb:ExtractMonad<'b, 'handle>) 
               (mc:ExtractMonad<'c, 'handle>) 
               (md:ExtractMonad<'d, 'handle>) 
               (me:ExtractMonad<'e, 'handle>) 
               (mf:ExtractMonad<'f, 'handle>) : ExtractMonad<'x, 'handle> = 
        ma >>= fun a ->
        mb >>= fun b ->
        mc >>= fun c ->
        md >>= fun d ->
        me >>= fun e ->
        mf >>= fun f ->
        mreturn (fn a b c d e f) 
        
    
    

    let tupleM2 (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'b, 'handle>) : ExtractMonad<'a * 'b, 'handle> = 
        liftM2 (fun a b -> (a,b)) ma mb

    let tupleM3 (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'b, 'handle>) 
                (mc:ExtractMonad<'c, 'handle>) : ExtractMonad<'a * 'b * 'c, 'handle> = 
        liftM3 (fun a b c -> (a,b,c)) ma mb mc

    let tupleM4 (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'b, 'handle>) 
                (mc:ExtractMonad<'c, 'handle>) 
                (md:ExtractMonad<'d, 'handle>) : ExtractMonad<'a * 'b * 'c * 'd, 'handle> = 
        liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

    let tupleM5 (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'b, 'handle>) 
                (mc:ExtractMonad<'c, 'handle>) 
                (md:ExtractMonad<'d, 'handle>) 
                (me:ExtractMonad<'e, 'handle>) : ExtractMonad<'a * 'b * 'c * 'd * 'e, 'handle> = 
        liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

    let tupleM6 (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'b, 'handle>) 
                (mc:ExtractMonad<'c, 'handle>) 
                (md:ExtractMonad<'d, 'handle>) 
                (me:ExtractMonad<'e, 'handle>) 
                (mf:ExtractMonad<'f, 'handle>) : ExtractMonad<'a * 'b * 'c * 'd * 'e * 'f, 'handle> = 
        liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf

    let pipeM2 (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'b, 'handle>) 
                (fn:'a -> 'b -> 'x) : ExtractMonad<'x, 'handle> = 
        liftM2 fn ma mb

    let pipeM3 (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'b, 'handle>) 
                (mc:ExtractMonad<'c, 'handle>) 
                (fn:'a -> 'b -> 'c -> 'x) : ExtractMonad<'x, 'handle> = 
        liftM3 fn ma mb mc

    let pipeM4 (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'b, 'handle>) 
                (mc:ExtractMonad<'c, 'handle>) 
                (md:ExtractMonad<'d, 'handle>) 
                (fn:'a -> 'b -> 'c -> 'd -> 'x) : ExtractMonad<'x, 'handle> = 
        liftM4 fn ma mb mc md

    let pipeM5 (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'b, 'handle>) 
                (mc:ExtractMonad<'c, 'handle>) 
                (md:ExtractMonad<'d, 'handle>) 
                (me:ExtractMonad<'e, 'handle>) 
                (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x) : ExtractMonad<'x, 'handle> = 
        liftM5 fn ma mb mc md me

    let pipeM6 (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'b, 'handle>) 
                (mc:ExtractMonad<'c, 'handle>) 
                (md:ExtractMonad<'d, 'handle>) 
                (me:ExtractMonad<'e, 'handle>) 
                (mf:ExtractMonad<'f, 'handle>) 
                (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) : ExtractMonad<'x, 'handle> = 
        liftM6 fn ma mb mc md me mf

        
    let many1M (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<'a list, 'handle> =
        pipeM2 ma (manyM ma) (fun x xs -> x :: xs)



    let skipMany1M (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<unit, 'handle> =
        pipeM2 ma (skipManyM ma) (fun _ _ -> ())

    let sepBy1M (ma:ExtractMonad<'a, 'handle>) 
                (msep:ExtractMonad<'sep, 'handle>) : ExtractMonad<'a list, 'handle> =
        pipeM2 ma (manyM (seqR msep ma)) (fun x xs -> x :: xs)

    let sepByM (ma:ExtractMonad<'a, 'handle>) 
                (msep:ExtractMonad<'sep, 'handle>) : ExtractMonad<'a list, 'handle> =
        altM (sepBy1M ma msep) (mreturn [])

    let endByM (ma:ExtractMonad<'a, 'handle>) 
                (mend:ExtractMonad<'sep, 'handle>) : ExtractMonad<'a list, 'handle> =
        seqL (manyM ma) mend

    let endBy1M (ma:ExtractMonad<'a, 'handle>) 
                (mend:ExtractMonad<'sep, 'handle>) : ExtractMonad<'a list, 'handle> =
        seqL (many1M ma) mend

    /// The end is optional...
    let sepEndByM (ma:ExtractMonad<'a, 'handle>) 
                    (msep:ExtractMonad<'sep, 'handle>) : ExtractMonad<'a list, 'handle> =
        seqL (sepByM ma msep) (ignoreM msep  <|> mreturn ())


    /// The end is optional...
    let sepEndBy1M (ma:ExtractMonad<'a, 'handle>) 
                    (msep:ExtractMonad<'sep, 'handle>) : ExtractMonad<'a list, 'handle> =
        seqL (sepBy1M ma msep) (ignoreM msep  <|> mreturn ())

