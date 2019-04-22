// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup

[<AutoOpen>]
module ExtractMonad = 
    
    
    open DocumentFormat.OpenXml
    open DocumentFormat.OpenXml.Packaging
    
    open DocSoup.Internal

    type ErrMsg = string

    type Answer<'a> = Result<'a, ErrMsg>

    type ExtractMonad<'handle, 'a> = 
        ExtractMonad of ('handle -> Answer<'a>)
        


    let inline private apply1 (ma: ExtractMonad<'handle, 'a>) 
                              (handle: 'handle) : Answer<'a>= 
        let (ExtractMonad f) = ma in f handle

        
    let inline mreturn (x:'a) : ExtractMonad<'handle, 'a> = 
        ExtractMonad <| fun _ -> Ok x

    let inline private bindM (ma:ExtractMonad<'handle, 'a>) 
        (f :'a -> ExtractMonad<'handle, 'b>) : ExtractMonad<'handle, 'b> =
            ExtractMonad <| fun handle -> 
                match apply1 ma handle with
                | Error msg -> Error msg
                | Ok a -> apply1 (f a) handle

    let inline private zeroM () : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun _ -> Error "zeroM"

    /// "First success"
    let inline private combineM (ma:ExtractMonad<'handle,'a>) 
                                (mb:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun handle -> 
            match apply1 ma handle with
            | Error msg -> apply1 mb handle
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

    let runExtractMonad (filePath:string) 
                        (project:WordprocessingDocument -> 'handle)  
                        (ma:ExtractMonad<'handle,'a>) : Result<'a,ErrMsg> = 
        match OpenXml.primitiveExtract filePath (fun doc -> apply1 ma (project doc)) with
        | Error msg -> Error msg
        | Ok ans -> ans


    // ****************************************************
    // Errors

    let throwError (msg:string) : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun _  -> Error msg

    let swapError (msg:string) (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun handle ->
            match apply1 ma handle with
            | Error _ -> Error msg
            | Ok a -> Ok a


    // ****************************************************
    // Focus

    let asks (project:'handle -> 'a) : ExtractMonad<'handle,'a> = 
        ExtractMonad <| fun handle -> 
            Ok (project handle)

    let local (focus:'handle1 -> 'handle2) 
              (ma:ExtractMonad<'handle2,'a>) : ExtractMonad<'handle1,'a> = 
        ExtractMonad <| fun handle -> 
            apply1 ma (focus handle)



    /// Chain a _selector_ and an _extractor_.
    let ( &>> ) (focus:ExtractMonad<'handle1, 'handle2>)
                (ma:ExtractMonad<'handle2,'ans>) : ExtractMonad<'handle1,'ans> = 
        bindM focus <| fun h1 -> local (fun _ -> h1) ma



    // ****************************************************
    // Monadic operations

    /// Bind operator
    let ( >>= ) (ma:ExtractMonad<'handle,'a>) 
              (fn:'a -> ExtractMonad<'handle,'b>) : ExtractMonad<'handle,'b> = 
        bindM ma fn

    /// Flipped Bind operator
    let ( =<< ) (fn:'a -> ExtractMonad<'handle,'b>) 
              (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'b> = 
        bindM ma fn


    let assertM (failMsg:string) (cond:ExtractMonad<'handle,bool>) : ExtractMonad<'handle,unit> = 
        ExtractMonad <| fun handle ->
            match apply1 cond handle with
            | Ok true -> Ok ()
            | _ -> Error failMsg
            
    let liftOption (opt:'a option) : ExtractMonad<'handle, 'a> = 
        match opt with
        | Some a -> mreturn a
        | None -> throwError "liftOption - None"

    /// fmap 
    let fmapM (fn:'a -> 'b) (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'b> = 
        ExtractMonad <| fun handle -> 
           match apply1 ma handle with
           | Error msg -> Error msg
           | Ok a -> Ok (fn a)

    /// Operator for fmap.
    let ( |>> ) (ma:ExtractMonad<'handle,'a>) (fn:'a -> 'b) : ExtractMonad<'handle,'b> = 
        fmapM fn ma

    /// Flipped fmap.
    let ( <<| ) (fn:'a -> 'b) (ma:ExtractMonad<'handle,'a>) : ExtractMonad<'handle,'b> = 
        fmapM fn ma


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



