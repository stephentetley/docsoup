// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module ExtractMonad = 
    
    open System.Text.RegularExpressions
    
    open DocumentFormat.OpenXml
    open DocumentFormat.OpenXml.Packaging
    
    open DocSoup.Internal

    type ErrMsg = string

    type Answer<'a> = Result<'a, ErrMsg>


    // To consider - maybe regex search options should go in the monad?
    type ExtractMonad<'handle, 'a> = 
        ExtractMonad of (RegexOptions -> 'handle -> Answer<'a>)
        


    let inline private apply1 (ma: ExtractMonad<'handle, 'a>)
                              (options: RegexOptions)
                              (handle: 'handle) : Answer<'a>= 
        let (ExtractMonad f) = ma in f options handle

        
    let inline mreturn (x:'a) : ExtractMonad<'handle, 'a> = 
        ExtractMonad <| fun _ _ -> Ok x

    let inline private bindM (ma:ExtractMonad<'handle, 'a>) 
        (f :'a -> ExtractMonad<'handle, 'b>) : ExtractMonad<'handle, 'b> =
            ExtractMonad <| fun opts handle -> 
                match apply1 ma opts handle with
                | Error msg -> Error msg
                | Ok a -> apply1 (f a) opts handle

    let inline private zeroM () : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun _ _ -> Error "zeroM"

    /// "First success"
    let inline private combineM (ma:ExtractMonad<'handle,'a>) 
                                (mb:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun opts handle -> 
            match apply1 ma opts handle with
            | Error msg -> apply1 mb opts handle
            | Ok a -> Ok a

    let inline private delayM (fn:unit -> ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'a> = 
        bindM (mreturn ()) fn 

    type ExtractMonadBuilder<'handle>() = 
        member self.Return(x:'a) : ExtractMonad<'handle, 'a>  = mreturn x
        member self.Bind (p:ExtractMonad<'handle,'a> , f: 'a -> ExtractMonad<'handle,'b>) : ExtractMonad<'handle,'b> = bindM p f
        member self.Zero () : ExtractMonad<'handle,'a> = zeroM ()
        member self.Combine (ma: ExtractMonad<'handle,'a>, mb: ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'a> = combineM ma mb
        member self.Delay (fn:unit -> ExtractMonad<'handle, 'a>) : ExtractMonad<'handle,'a> = delayM fn
        member self.ReturnFrom(ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle, 'a> = ma

    // ****************************************************
    // Run

    let private isWordTempFile (filePath:string) : bool = 
        let fileName = System.IO.FileInfo(filePath).Name
        Regex.IsMatch(input = fileName, pattern = "^~\$")

    let runExtractMonad (filePath:string) 
                        (project:WordprocessingDocument -> 'handle)  
                        (ma:ExtractMonad<'handle,'a>) : Result<'a,ErrMsg> = 
        let opts = RegexOptions.None
        if isWordTempFile filePath then 
            Error (sprintf "Invalid Word file: %s" filePath) 
        else
            match OpenXml.primitiveExtract filePath 
                                           (fun doc -> apply1 ma opts (project doc)) with
            | Error msg -> Error msg
            | Ok ans -> ans


    // ****************************************************
    // Errors

    let extractError (msg:string) : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun _ _ -> Error msg

    let swapError (msg:string) (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun opts handle ->
            match apply1 ma opts handle with
            | Error _ -> Error msg
            | Ok a -> Ok a


    // ****************************************************
    // Bind operations

    /// Bind operator
    let ( >>= ) (ma:ExtractMonad<'handle,'a>) 
              (fn:'a -> ExtractMonad<'handle,'b>) : ExtractMonad<'handle,'b> = 
        bindM ma fn

    /// Flipped Bind operator
    let ( =<< ) (fn:'a -> ExtractMonad<'handle,'b>) 
              (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'b> = 
        bindM ma fn

    // ****************************************************
    // Focus

    let asks (project:'handle -> 'a) : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun _ handle -> 
            Ok (project handle)

    let local (adaptor:'handle1 -> 'handle2) 
              (ma:ExtractMonad<'handle2,'a>) : ExtractMonad<'handle1,'a> = 
        ExtractMonad <| fun opts handle -> 
            apply1 ma opts (adaptor handle)

    let focus (newFocus:'handle2) 
              (ma:ExtractMonad<'handle2,'a>) : ExtractMonad<'handle1,'a> = 
        ExtractMonad <| fun opts _ -> 
            apply1 ma opts newFocus

    /// Chain a _selector_ and an _extractor_.
    let focusM (selector:ExtractMonad<'handle1, 'handle2>)
               (ma:ExtractMonad<'handle2,'ans>) : ExtractMonad<'handle1,'ans> = 
        selector >>= fun newHandle -> 
        focus newHandle ma

    /// Operator for focusM.
    /// Chain a _selector_ and an _extractor_.
    let ( &>> ) (selector:ExtractMonad<'handle1, 'handle2>)
                (ma:ExtractMonad<'handle2,'ans>) : ExtractMonad<'handle1,'ans> = 
        focusM selector ma


    // ****************************************************
    // Monadic operations

    let getRegexOptions () : ExtractMonad<'handle, RegexOptions> = 
        ExtractMonad <| fun opts _ -> Ok opts

    let localOptions (modify : RegexOptions -> RegexOptions)
                     (ma: ExtractMonad<'handle, 'a>) : ExtractMonad<'handle, 'a> = 
        ExtractMonad <| fun opts handle -> 
            apply1 ma (modify opts) handle

    let ignoreCase (ma: ExtractMonad<'handle, 'a>) : ExtractMonad<'handle, 'a> = 
        localOptions (fun opts -> RegexOptions.IgnoreCase ||| opts) ma

    // ****************************************************
    // Monadic operations


    /// fmap 
    let fmapM (fn:'a -> 'b) (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'b> = 
        ExtractMonad <| fun opts handle -> 
           match apply1 ma opts handle with
           | Error msg -> Error msg
           | Ok a -> Ok (fn a)

    /// Operator for fmap.
    let ( |>> ) (ma:ExtractMonad<'handle,'a>) (fn:'a -> 'b) : ExtractMonad<'handle,'b> = 
        fmapM fn ma

    /// Flipped fmap.
    let ( <<| ) (fn:'a -> 'b) (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'b> = 
        fmapM fn ma

    /// Perform an action, but ignore its answer.
    let ignoreM (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle, unit> = 
        ma |>> fun _ -> ()

    let assertM (failMsg:string) (cond:ExtractMonad<'handle,bool>) : ExtractMonad<'handle,unit> = 
        ExtractMonad <| fun opts handle ->
            match apply1 cond opts handle with
            | Ok true -> Ok ()
            | _ -> Error failMsg

    /// Lift an option value.
    /// Some ans becomes a success value in the ExtractMonad.
    /// None throws a monadic error (not system system exception). 
    let liftOption (opt:'a option) : ExtractMonad<'handle, 'a> = 
        match opt with
        | Some a -> mreturn a
        | None -> extractError "liftOption - None"

    /// Lift an action that may fail (e.g. an IO operation).
    /// If the action does fail, replace the hard error with 
    /// a (soft) error within the monad.
    let liftAction (errMsg:string) (action: unit -> 'a) : ExtractMonad<'handle, 'a> = 
        try 
            let ans = action ()
            mreturn ans
        with
        | _ -> extractError errMsg


    // liftM (which is fmap)
    let liftM (fn:'a -> 'x) (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'x> = 
        fmapM fn ma

    let liftM2 (fn:'a -> 'b -> 'x) 
               (ma:ExtractMonad<'handle,'a>) 
               (mb:ExtractMonad<'handle,'b>) : ExtractMonad<'handle,'x> = 
        ma >>= fun a ->
        mb >>= fun b ->
        mreturn (fn a b)

    let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
               (ma:ExtractMonad<'handle,'a>) 
               (mb:ExtractMonad<'handle,'b>) 
               (mc:ExtractMonad<'handle,'c>) : ExtractMonad<'handle,'x> = 
        ma >>= fun a ->
        mb >>= fun b ->
        mc >>= fun c ->
        mreturn (fn a b c)

    let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
               (ma:ExtractMonad<'handle,'a>) 
               (mb:ExtractMonad<'handle,'b>) 
               (mc:ExtractMonad<'handle,'c>) 
               (md:ExtractMonad<'handle,'d>) : ExtractMonad<'handle,'x> = 
        ma >>= fun a ->
        mb >>= fun b ->
        mc >>= fun c ->
        md >>= fun d ->
        mreturn (fn a b c d)


    let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) 
               (ma:ExtractMonad<'handle,'a>) 
               (mb:ExtractMonad<'handle,'b>) 
               (mc:ExtractMonad<'handle,'c>) 
               (md:ExtractMonad<'handle,'d>) 
               (me:ExtractMonad<'handle,'e>) : ExtractMonad<'handle,'x> = 
        ma >>= fun a ->
        mb >>= fun b ->
        mc >>= fun c ->
        md >>= fun d ->
        me >>= fun e ->
        mreturn (fn a b c d e)
        

    let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
               (ma:ExtractMonad<'handle,'a>) 
               (mb:ExtractMonad<'handle,'b>) 
               (mc:ExtractMonad<'handle,'c>) 
               (md:ExtractMonad<'handle,'d>) 
               (me:ExtractMonad<'handle,'e>) 
               (mf:ExtractMonad<'handle,'f>) : ExtractMonad<'handle,'x> = 
        ma >>= fun a ->
        mb >>= fun b ->
        mc >>= fun c ->
        md >>= fun d ->
        me >>= fun e ->
        mf >>= fun f ->
        mreturn (fn a b c d e f)    

    let tupleM2 (ma:ExtractMonad<'handle,'a>) 
                (mb:ExtractMonad<'handle,'b>) : ExtractMonad<'handle,'a * 'b> = 
        liftM2 (fun a b -> (a,b)) ma mb

    let tupleM3 (ma:ExtractMonad<'handle,'a>) 
                (mb:ExtractMonad<'handle,'b>) 
                (mc:ExtractMonad<'handle,'c>) : ExtractMonad<'handle,'a * 'b * 'c> = 
        liftM3 (fun a b c -> (a,b,c)) ma mb mc

    let tupleM4 (ma:ExtractMonad<'handle,'a>) 
                (mb:ExtractMonad<'handle,'b>) 
                (mc:ExtractMonad<'handle,'c>) 
                (md:ExtractMonad<'handle,'d>) : ExtractMonad<'handle,'a * 'b * 'c * 'd> = 
        liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

    let tupleM5 (ma:ExtractMonad<'handle,'a>) 
                (mb:ExtractMonad<'handle,'b>) 
                (mc:ExtractMonad<'handle,'c>) 
                (md:ExtractMonad<'handle,'d>) 
                (me:ExtractMonad<'handle,'e>) : ExtractMonad<'handle,'a * 'b * 'c * 'd * 'e> = 
        liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

    let tupleM6 (ma:ExtractMonad<'handle,'a>) 
                (mb:ExtractMonad<'handle,'b>) 
                (mc:ExtractMonad<'handle,'c>) 
                (md:ExtractMonad<'handle,'d>) 
                (me:ExtractMonad<'handle,'e>) 
                (mf:ExtractMonad<'handle,'f>) : ExtractMonad<'handle,'a * 'b * 'c * 'd * 'e * 'f> = 
        liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf

    let pipeM2 (ma:ExtractMonad<'handle,'a>) 
               (mb:ExtractMonad<'handle,'b>) 
               (fn:'a -> 'b -> 'x) : ExtractMonad<'handle,'x> = 
        liftM2 fn ma mb

    let pipeM3 (ma:ExtractMonad<'handle,'a>) 
               (mb:ExtractMonad<'handle,'b>) 
               (mc:ExtractMonad<'handle,'c>) 
               (fn:'a -> 'b -> 'c -> 'x) : ExtractMonad<'handle,'x> = 
        liftM3 fn ma mb mc

    let pipeM4 (ma:ExtractMonad<'handle,'a>) 
               (mb:ExtractMonad<'handle,'b>) 
               (mc:ExtractMonad<'handle,'c>) 
               (md:ExtractMonad<'handle,'d>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'x) : ExtractMonad<'handle,'x> = 
        liftM4 fn ma mb mc md

    let pipeM5 (ma:ExtractMonad<'handle,'a>) 
               (mb:ExtractMonad<'handle,'b>) 
               (mc:ExtractMonad<'handle,'c>) 
               (md:ExtractMonad<'handle,'d>) 
               (me:ExtractMonad<'handle,'e>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x) : ExtractMonad<'handle,'x> = 
        liftM5 fn ma mb mc md me

    let pipeM6 (ma:ExtractMonad<'handle,'a>) 
               (mb:ExtractMonad<'handle,'b>) 
               (mc:ExtractMonad<'handle,'c>) 
               (md:ExtractMonad<'handle,'d>) 
               (me:ExtractMonad<'handle,'e>) 
               (mf:ExtractMonad<'handle,'f>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) : ExtractMonad<'handle,'x> = 
        liftM6 fn ma mb mc md me mf

        
    /// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
    let altM (ma:ExtractMonad<'handle,'a>) (mb:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'a> = 
        combineM ma mb


    /// Alt operator
    let ( <|> ) (ma:ExtractMonad<'handle,'a>) 
                (mb:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'a> = 
        combineM ma mb


    /// Haskell Applicative's (<*>)
    let apM (mf:ExtractMonad<'handle,'a ->'b>) (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'b> = 
        mf >>= fun fn -> 
        ma >>= fun a  -> 
        mreturn (fn a)

    /// Perform two actions in sequence. 
    /// Ignore the results of the second action if both succeed.
    let seqL (ma:ExtractMonad<'handle,'a>) (mb:ExtractMonad<'handle,'b>) : ExtractMonad<'handle,'a> = 
        ma >>= fun ans -> 
        mb >>= fun _ -> 
        mreturn ans

    /// Perform two actions in sequence. 
    /// Ignore the results of the first action if both succeed.
    let seqR (ma:ExtractMonad<'handle,'a>) (mb:ExtractMonad<'handle,'b>) : ExtractMonad<'handle,'b> = 
        ma >>= fun _ -> 
        mb >>= fun ans -> 
        mreturn ans

    /// Operator for seqL
    let (.>>) (ma:ExtractMonad<'handle,'a>) 
              (mb:ExtractMonad<'handle,'b>) : ExtractMonad<'handle,'a> = 
        seqL ma mb

    /// Operator for seqR
    let (>>.) (ma:ExtractMonad<'handle,'a>) 
              (mb:ExtractMonad<'handle,'b>) : ExtractMonad<'handle,'b> = 
        seqR ma mb


    let kleisliL (mf:'a -> ExtractMonad<'handle,'b>)
                 (mg:'b -> ExtractMonad<'handle,'c>)
                 (source:'a) : ExtractMonad<'handle,'c> = 
        mf source >>= fun a1 ->
        mg a1 >>= fun ans ->
        mreturn ans


    /// Flipped kleisliL
    let kleisliR (mf:'b -> ExtractMonad<'handle,'c>)
                 (mg:'a -> ExtractMonad<'handle,'b>)
                 (source:'a) : ExtractMonad<'handle,'c> = 
        kleisliL mg mf source



    /// Operator for kleisliL
    let (>=>) (mf : 'a -> ExtractMonad<'handle,'b>)
              (mg : 'b -> ExtractMonad<'handle,'c>)
              (source:'a) : ExtractMonad<'handle,'c> = 
        kleisliL mf mg source


    /// Operator for kleisliR
    let (<=<) (mf : 'b -> ExtractMonad<'handle,'c>)
              (mg : 'a -> ExtractMonad<'handle,'b>)
              (source:'a) : ExtractMonad<'handle,'c> = 
        kleisliR mf mg source


    let andM (ma:ExtractMonad<'handle,bool>)
             (mb:ExtractMonad<'handle,bool>) : ExtractMonad<'handle,bool> =
        pipeM2 ma mb (&&)


    let ( <&&> ) (ma:ExtractMonad<'handle,bool>)
                 (mb:ExtractMonad<'handle,bool>) : ExtractMonad<'handle,bool> =
        andM ma mb


    let orM (ma:ExtractMonad<'handle,bool>)
            (mb:ExtractMonad<'handle,bool>) : ExtractMonad<'handle,bool> =
        ma >>= fun ans -> 
        if ans then 
            mreturn true 
        else 
            mb >>= fun ans2 -> mreturn ans2

    let ( <||> ) (ma:ExtractMonad<'handle,bool>)
                 (mb:ExtractMonad<'handle,bool>) : ExtractMonad<'handle,bool> =
        orM ma mb

    // ****************************************************
    // List processing

    /// This interprets failure as false.
    /// Is this correct... wise...
    let allM (predicates: ExtractMonad<'handle, bool> list) : ExtractMonad<'handle,bool> =
        ExtractMonad <| fun opts handle -> 
            let rec work ys cont = 
                match ys with
                | [] -> cont (Ok true)
                | test :: rest -> 
                    match apply1 test opts handle with
                    | Error msg -> cont (Ok false)    // short circuit
                    | Ok false -> cont (Ok false)    // short circuit
                    | Ok true -> work rest cont
            work predicates (fun ans -> ans)

    /// This interprets failure as false.
    /// Is this correct... wise...
    let anyM (predicates: ExtractMonad<'handle, bool> list) : ExtractMonad<'handle,bool> =
        ExtractMonad <| fun opts handle -> 
            let rec work ys cont = 
                match ys with
                | [] -> cont (Ok false)
                | test :: rest -> 
                    match apply1 test opts handle with
                    | Error msg -> work rest cont   
                    | Ok false -> work rest cont    // short circuit
                    | Ok true -> cont (Ok true)
            work predicates (fun ans -> ans)


    /// Implemented in CPS 
    let mapM (mf: 'a -> ExtractMonad<'handle,'b>) 
             (source:'a list) : ExtractMonad<'handle,'b list> = 
        ExtractMonad <| fun opts handle -> 
            let rec work ys fk sk = 
                match ys with
                | [] -> sk []
                | z :: zs -> 
                    match apply1 (mf z) opts handle with
                    | Error msg -> fk msg
                    | Ok ans -> 
                        work zs fk (fun accs ->
                        sk (ans::accs))
            work source (fun msg -> Error msg) (fun ans -> Ok ans)

    /// Flipped mapM
    let forM (source:'a list) 
             (mf: 'a -> ExtractMonad<'handle,'b>) : ExtractMonad<'handle,'b list> = 
        mapM mf source


    /// Implemented in CPS 
    let mapMz (mf: 'a -> ExtractMonad<'handle,'b>) 
             (source:'a list) : ExtractMonad<'handle,unit> = 
        ExtractMonad <| fun opts handle -> 
            let rec work ys cont = 
                match ys with
                | [] -> cont (Ok ())
                | z :: zs -> 
                    match apply1 (mf z) opts handle with
                    | Error msg -> cont (Error msg)
                    | Ok _ -> 
                        work zs cont
            work source (fun ans -> ans)

    /// Flipped mapM
    let forMz (source:'a list) 
             (mf: 'a -> ExtractMonad<'handle,'b>) : ExtractMonad<'handle,unit> = 
        mapMz mf source


    let manyM (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle, 'a list> =
        ExtractMonad <| fun opts handle -> 
            let rec work cont = 
                match apply1 ma opts handle with
                | Error msg -> cont []
                | Ok ans -> 
                    work (fun acc -> cont (ans::acc))
            work (fun ans -> Ok ans)

    let many1M (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle, 'a list> =
        pipeM2 ma (manyM ma) (fun x xs -> x :: xs)



    let skipManyM (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle, unit> =
        ExtractMonad <| fun opts handle -> 
            let rec work cont = 
                match apply1 ma opts handle with
                | Error msg -> cont ()
                | Ok ans -> 
                    work cont
            work (fun ans -> Ok ans)

    let skipMany1M (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle, unit> =
        pipeM2 ma (skipManyM ma) (fun _ _ -> ())

    let sepBy1M (ma:ExtractMonad<'handle,'a>) 
                (msep:ExtractMonad<'handle,'sep>) : ExtractMonad<'handle, 'a list> =
        pipeM2 ma (manyM (seqR msep ma)) (fun x xs -> x :: xs)

    let sepByM (ma:ExtractMonad<'handle,'a>) 
               (msep:ExtractMonad<'handle,'sep>) : ExtractMonad<'handle, 'a list> =
        altM (sepBy1M ma msep) (mreturn [])

    let endByM (ma:ExtractMonad<'handle,'a>) 
               (mend:ExtractMonad<'handle,'sep>) : ExtractMonad<'handle, 'a list> =
        seqL (manyM ma) mend

    let endBy1M (ma:ExtractMonad<'handle,'a>) 
                (mend:ExtractMonad<'handle,'sep>) : ExtractMonad<'handle, 'a list> =
        seqL (many1M ma) mend

    /// The end is optional...
    let sepEndByM (ma:ExtractMonad<'handle,'a>) 
                  (msep:ExtractMonad<'handle,'sep>) : ExtractMonad<'handle, 'a list> =
        seqL (sepByM ma msep) (ignoreM msep  <|> mreturn ())


    /// The end is optional...
    let sepEndBy1M (ma:ExtractMonad<'handle,'a>) 
                   (msep:ExtractMonad<'handle,'sep>) : ExtractMonad<'handle, 'a list> =
        seqL (sepBy1M ma msep) (ignoreM msep  <|> mreturn ())

    let countM (ntimes:int) 
               (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle, 'a list> =
        ExtractMonad <| fun opts handle -> 
            let rec work ix fk sk = 
                if ix <= 0 then 
                    sk []
                else
                    match apply1 ma opts handle with
                    | Error msg -> fk (sprintf "count - %s msg" msg)
                    | Ok ans -> 
                        work (ix-1) fk (fun acc -> sk (ans::acc))
            work ntimes (fun msg -> Error msg) (fun ans -> Ok ans)


    let findM (predicate:'a -> ExtractMonad<'handle,bool>)
              (items:'a list) : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun opts handle -> 
            let rec work xs fk sk = 
                match xs with
                | [] -> fk "findM not found"
                | item :: rest -> 
                    match apply1 (predicate item) opts handle with
                    | Ok true -> 
                        sk item
                    | Ok false -> work rest fk sk 
                    | Error msg -> work rest fk sk 
            work items (fun msg -> Error msg) (fun ans -> Ok ans)

    let findIndexM (predicate:'a -> ExtractMonad<'handle,bool>)
                   (items:'a list) : ExtractMonad<'handle,int> = 
        ExtractMonad <| fun opts handle -> 
            let rec work ix xs fk sk = 
                match xs with
                | [] -> fk "findM not found"
                | item :: rest -> 
                    match apply1 (predicate item) opts handle with
                    | Ok true -> 
                        sk ix
                    | Ok false -> work (ix+1) rest  fk sk 
                    | Error msg -> work (ix+1) rest fk sk 
            work 0 items (fun msg -> Error msg) (fun ans -> Ok ans)

    let forallM (predicate:'a -> ExtractMonad<'handle,bool>)
                (items:'a list) : ExtractMonad<'handle, bool> = 
        ExtractMonad <| fun opts handle -> 
            let rec work xs cont = 
                match xs with
                | [] -> cont true
                | item :: rest -> 
                    match apply1 (predicate item) opts handle with
                    | Ok true -> work rest cont
                    | Ok false -> cont false
                    | Error msg -> cont false
            work items (fun ans -> Ok ans)

    /// Note we have "natural failure" in the ExtractMonad so 
    /// pickM does not have to return an Option.
    let pickM (chooser:'a -> ExtractMonad<'handle,'b>)
              (items:'a list) : ExtractMonad<'handle, 'b> = 
        ExtractMonad <| fun opts handle -> 
            let rec work xs fk sk = 
                match xs with
                | [] -> fk "pickM not found"
                | item :: rest -> 
                    match apply1 (chooser item) opts handle with
                    | Ok ans -> sk ans
                    | Error _ -> work rest fk sk
            work items (fun msg -> Error msg) (fun ans -> Ok ans)

