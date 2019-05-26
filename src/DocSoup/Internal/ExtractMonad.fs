// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup.Internal


module ExtractMonad = 
    
    open System.Text.RegularExpressions
    
    open DocumentFormat.OpenXml
    open DocumentFormat.OpenXml.Packaging
    
    open DocSoup.Internal

    type ErrMsg = string

    type State = int

    type Answer<'a> = Result<'a * State, ErrMsg>


    // To consider - maybe regex search options should go in the monad?
    type ExtractMonad<'a, 'handle> = 
        ExtractMonad of (RegexOptions -> 'handle -> State -> Answer<'a>)
        


    let inline private apply1 (ma: ExtractMonad<'a, 'handle>)
                              (options: RegexOptions)
                              (handle: 'handle) 
                              (state:State) : Answer<'a>= 
        let (ExtractMonad f) = ma in f options handle state

        
    let inline mreturn (x:'a) : ExtractMonad<'a, 'handle> = 
        ExtractMonad <| fun _ _ st -> Ok (x, st)

    let inline private bindM (ma:ExtractMonad<'a, 'handle>) 
        (f :'a -> ExtractMonad<'b, 'handle>) : ExtractMonad<'b, 'handle> =
            ExtractMonad <| fun opts handle state -> 
                match apply1 ma opts handle state with
                | Error msg -> Error msg
                | Ok (a, st1) -> apply1 (f a) opts handle st1

    let inline private zeroM () : ExtractMonad<'a, 'handle> = 
        ExtractMonad <| fun _ _ _ -> Error "zeroM"

    /// "First success"
    let inline private combineM (ma:ExtractMonad<'a, 'handle>) 
                                (mb:ExtractMonad<'a, 'handle>) : ExtractMonad<'a, 'handle> = 
        ExtractMonad <| fun opts handle state -> 
            match apply1 ma opts handle state with
            | Error msg -> apply1 mb opts handle state
            | Ok a -> Ok a

    let inline private delayM (fn:unit -> ExtractMonad<'a, 'handle>) : ExtractMonad<'a, 'handle> = 
        bindM (mreturn ()) fn 

    type ExtractMonadBuilder<'handle>() = 
        member self.Return(x:'a) : ExtractMonad<'a, 'handle>  = mreturn x
        member self.Bind (p:ExtractMonad<'a, 'handle> , f: 'a -> ExtractMonad<'b, 'handle>) : ExtractMonad<'b, 'handle> = bindM p f
        member self.Zero () : ExtractMonad<'a, 'handle> = zeroM ()
        member self.Combine (ma: ExtractMonad<'a, 'handle>, mb: ExtractMonad<'a, 'handle>) : ExtractMonad<'a, 'handle> = combineM ma mb
        member self.Delay (fn:unit -> ExtractMonad<'a, 'handle>) : ExtractMonad<'a, 'handle> = delayM fn
        member self.ReturnFrom(ma:ExtractMonad<'a, 'handle>) : ExtractMonad<'a, 'handle> = ma

    // ****************************************************
    // Run

    let private isWordTempFile (filePath:string) : bool = 
        let fileName = System.IO.FileInfo(filePath).Name
        Regex.IsMatch(input = fileName, pattern = "^~\$")


    let runExtractMonad (filePath:string) 
                        (project:WordprocessingDocument -> 'handle)  
                        (ma:ExtractMonad<'a, 'handle>) : Result<'a, ErrMsg> = 
        let opts = RegexOptions.None
        if isWordTempFile filePath then 
            Error (sprintf "Invalid Word file: %s" filePath) 
        else
            match OpenXml.primitiveExtract filePath 
                                           (fun doc -> apply1 ma opts (project doc) 0) with
            | Error msg -> Error msg
            | Ok (Ok (ans, _)) -> Ok ans
            | Ok (Error msg) -> Error msg


    let internalRunExtract (handle:'handle)  
                           (ma:ExtractMonad<'a, 'handle>) : Result<'a,ErrMsg> = 
        let opts = RegexOptions.None
        match apply1 ma opts handle 0 with
        | Error msg -> Error msg
        | Ok (ans, _) -> Ok ans

    // ****************************************************
    // Errors

    let extractError (msg:string) : ExtractMonad<'a, 'handle> = 
        ExtractMonad <| fun _ _ _ -> Error msg

    let swapError (msg:string) (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<'a, 'handle> = 
        ExtractMonad <| fun opts handle state ->
            match apply1 ma opts handle state with
            | Error _ -> Error msg
            | Ok a -> Ok a




    // ****************************************************
    // Bind operations

    /// Bind operator
    let ( >>= ) (ma:ExtractMonad<'a, 'handle>) 
              (fn:'a -> ExtractMonad<'b, 'handle>) : ExtractMonad<'b, 'handle> = 
        bindM ma fn

    /// Flipped Bind operator
    let ( =<< ) (fn:'a -> ExtractMonad<'b, 'handle>) 
              (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<'b, 'handle> = 
        bindM ma fn

    // ****************************************************
    // Focus

    let asks (project:'handle -> 'a) : ExtractMonad<'a, 'handle> = 
        ExtractMonad <| fun _ handle state -> 
            Ok (project handle, state)

    let local (adaptor:'handle1 -> 'handle2) 
              (ma:ExtractMonad<'a, 'handle2>) : ExtractMonad<'a, 'handle1> = 
        ExtractMonad <| fun opts handle state -> 
            apply1 ma opts (adaptor handle) state

    let focus (newFocus:'handle2) 
              (ma:ExtractMonad<'a, 'handle2>) : ExtractMonad<'a, 'handle1> = 
        ExtractMonad <| fun opts _ -> 
            apply1 ma opts newFocus

    /// Chain a _selector_ and an _extractor_.
    let focusM (selector:ExtractMonad<'handle2, 'handle1>)
               (ma:ExtractMonad<'a, 'handle2>) : ExtractMonad<'a, 'handle1> = 
        selector >>= fun newHandle -> 
        focus newHandle ma


    // ****************************************************
    // State operations

    let getPosition () : ExtractMonad<int, 'handle> = 
        ExtractMonad <| fun opts handle state -> 
            Ok (state, state)


    let incrPosition () : ExtractMonad<unit, 'handle> = 
        ExtractMonad <| fun opts handle state -> 
            let state1 = state + 1 in Ok ((), state1)

    let peek (errMsg:ErrMsg) (getter: State -> 'handle -> 'ans) : ExtractMonad<'ans, 'handle> = 
        ExtractMonad <| fun _ handle state -> 
            try 
                let ans = getter state handle 
                Ok (ans, state)
            with
            | _ -> Error errMsg

    let consume1 (errMsg:ErrMsg) (getter:State -> 'handle -> 'ans) : ExtractMonad<'ans, 'handle> = 
        peek errMsg getter >>= fun ans -> 
        incrPosition () >>= fun _ ->
        mreturn ans



    // ****************************************************
    // Regex options

    let getRegexOptions () : ExtractMonad<RegexOptions, 'handle> = 
        ExtractMonad <| fun opts _ state -> Ok (opts, state)

    let localOptions (modify : RegexOptions -> RegexOptions)
                     (ma: ExtractMonad<'a, 'handle>) : ExtractMonad<'a, 'handle> = 
        ExtractMonad <| fun opts handle state-> 
            apply1 ma (modify opts) handle state



    // ****************************************************
    // Monadic operations


    /// fmap 
    let fmapM (fn:'a -> 'b) (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<'b, 'handle> = 
        ExtractMonad <| fun opts handle state-> 
           match apply1 ma opts handle state with
           | Error msg -> Error msg
           | Ok (a, st1) -> Ok (fn a, st1)




        
    /// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
    let altM (ma:ExtractMonad<'a, 'handle>) (mb:ExtractMonad<'a, 'handle>) : ExtractMonad<'a, 'handle> = 
        combineM ma mb


    /// Alt operator
    let ( <|> ) (ma:ExtractMonad<'a, 'handle>) 
                (mb:ExtractMonad<'a, 'handle>) : ExtractMonad<'a, 'handle> = 
        combineM ma mb



    // ****************************************************
    // List processing

    /// This pushes up failure of a parse as failure of allM.
    let allM (predicates: ExtractMonad<bool, 'handle> list) : ExtractMonad<bool, 'handle> =
        ExtractMonad <| fun opts handle state0 -> 
            let rec work ys st cont = 
                match ys with
                | [] -> cont (Ok (true, st))
                | test :: rest -> 
                    match apply1 test opts handle st with
                    | Error msg -> cont (Error msg)    // short circuit
                    | Ok (false, st1) -> cont (Ok (false, st1))    // short circuit
                    | Ok (true, st1) -> work rest st1 cont
            work predicates state0 (fun ans -> ans)

    /// This pushes up failure of a parse as failure of anyM.
    let anyM (predicates: ExtractMonad<bool, 'handle> list) : ExtractMonad<bool, 'handle> =
        ExtractMonad <| fun opts handle state0 -> 
            let rec work ys st cont = 
                match ys with
                | [] -> cont (Ok (false, st))
                | test :: rest -> 
                    match apply1 test opts handle st with
                    | Error msg -> cont (Error msg)             // short circuit
                    | Ok (false, st1) -> work rest st cont
                    | Ok (true, st1) -> cont (Ok (true, st1))    // short circuit
            work predicates state0 (fun ans -> ans)


    /// Implemented in CPS 
    let mapM (mf: 'a -> ExtractMonad<'b, 'handle>) 
             (source:'a list) : ExtractMonad<'b list, 'handle> = 
        ExtractMonad <| fun opts handle state0 -> 
            let rec work ys st fk sk = 
                match ys with
                | [] -> sk st []
                | z :: zs -> 
                    match apply1 (mf z) opts handle st with
                    | Error msg -> fk msg
                    | Ok (ans, st1) -> 
                        work zs st1 fk (fun st2 accs ->
                        sk st2 (ans::accs))
            work source state0 (fun msg -> Error msg) (fun st ans -> Ok (ans, st))

    /// Flipped mapM
    let forM (source:'a list) 
             (mf: 'a -> ExtractMonad<'b, 'handle>) : ExtractMonad<'b list, 'handle> = 
        mapM mf source


    /// Implemented in CPS 
    let mapMz (mf: 'a -> ExtractMonad<'b, 'handle>) 
              (source:'a list) : ExtractMonad<unit, 'handle> = 
        ExtractMonad <| fun opts handle state0 -> 
            let rec work ys st cont = 
                match ys with
                | [] -> cont (Ok ((), st))
                | z :: zs -> 
                    match apply1 (mf z) opts handle st with
                    | Error msg -> cont (Error msg)
                    | Ok (_, st1) -> 
                        work zs st1 cont
            work source state0 (fun ans -> ans)

    /// Flipped mapM
    let forMz (source: 'a list) 
              (mf: 'a -> ExtractMonad<'b, 'handle>) : ExtractMonad<unit, 'handle> = 
        mapMz mf source

    
    /// CAUTION - without a notion of progress, this parser does not make sense.
    let manyM (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<'a list, 'handle> =
        ExtractMonad <| fun opts handle state0 -> 
            let rec work st cont = 
                match apply1 ma opts handle st with
                | Error msg -> cont st []
                | Ok (ans, st1) -> 
                    work st1 (fun st2 acc -> cont st2 (ans::acc))
            work state0 (fun st ans -> Ok (ans, st))



    let skipManyM (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<unit, 'handle> =
        ExtractMonad <| fun opts handle state0 -> 
            let rec work st cont = 
                match apply1 ma opts handle st with
                | Error msg -> cont (Ok ((), st))
                | Ok (ans, st1) -> 
                    work st1 cont
            work state0 (fun ans -> ans)

    let countM (ntimes:int) 
               (ma:ExtractMonad<'a, 'handle>) : ExtractMonad<'a list, 'handle> =
        ExtractMonad <| fun opts handle state0 -> 
            let rec work ix st fk sk = 
                if ix <= 0 then 
                    sk st []
                else
                    match apply1 ma opts handle st with
                    | Error msg -> fk (sprintf "count - %s msg" msg)
                    | Ok (ans, st1) -> 
                        work (ix-1) st1 fk (fun st2 acc -> sk st2 (ans::acc))
            work ntimes state0 (fun msg -> Error msg) (fun st ans -> Ok (ans, st))


    let findM (predicate:'a -> ExtractMonad<bool, 'handle>)
              (items:'a list) : ExtractMonad<'a, 'handle> = 
        ExtractMonad <| fun opts handle state0 -> 
            let rec work xs st fk sk = 
                match xs with
                | [] -> fk "findM not found"
                | item :: rest -> 
                    match apply1 (predicate item) opts handle st with
                    | Ok (true, st1) -> sk st1 item
                    | Ok (false, st1) -> work rest st1 fk sk 
                    | Error msg -> fk msg
            work items state0 (fun msg -> Error msg) (fun st ans -> Ok (ans, st))

    let findIndexM (predicate:'a -> ExtractMonad<bool, 'handle>)
                   (items:'a list) : ExtractMonad<int, 'handle> = 
        ExtractMonad <| fun opts handle state0-> 
            let rec work ix xs st fk sk = 
                match xs with
                | [] -> fk "findM not found"
                | item :: rest -> 
                    match apply1 (predicate item) opts handle st with
                    | Ok (true, st1) -> sk st1 ix
                    | Ok (false, st1) -> work (ix+1) rest st1 fk sk 
                    | Error msg -> fk msg 
            work 0 items state0 (fun msg -> Error msg) (fun st ans -> Ok (ans, st))

    let forallM (predicate:'a -> ExtractMonad<bool, 'handle>)
                (items:'a list) : ExtractMonad<bool, 'handle> = 
        ExtractMonad <| fun opts handle state0 -> 
            let rec work xs st fk sk = 
                match xs with
                | [] -> sk st true
                | item :: rest -> 
                    match apply1 (predicate item) opts handle st with
                    | Ok (true, st1) -> work rest st1 fk sk
                    | Ok (false, st1) -> sk st1 false
                    | Error msg -> fk msg
            work items state0 (fun msg -> Error msg) (fun st ans -> Ok (ans, st))

    /// Note we have "natural failure" in the ExtractMonad so 
    /// pickM does not have to return an Option.
    let pickM (chooser:'a -> ExtractMonad<'b, 'handle>)
              (items:'a list) : ExtractMonad<'b, 'handle> = 
        ExtractMonad <| fun opts handle state0 -> 
            let rec work xs st fk sk = 
                match xs with
                | [] -> fk "pickM not found"
                | item :: rest -> 
                    match apply1 (chooser item) opts handle st with
                    | Ok (ans, st1) -> sk st1 ans
                    | Error _ -> work rest st fk sk
            work items state0 (fun msg -> Error msg) (fun st ans -> Ok (ans, st))

