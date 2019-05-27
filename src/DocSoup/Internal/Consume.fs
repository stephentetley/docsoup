// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocSoup.Internal


module Consume = 
    
    open DocSoup.Internal.ExtractMonad

    /// NOTE - not sure if a "first class module" is really required.

    type ConsumeModule<'element> = 
        { EndOfInput : ExtractMonad<unit, 'element []>
          GetItem : ExtractMonad<'element, 'element []>
          GetItems : int -> ExtractMonad<'element [], 'element []>
          Position : ExtractMonad<int, 'element []>
          GetInput : ExtractMonad<'element [], 'element []>
          InputCount : ExtractMonad<int, 'element []>
          Satisfy : ('element -> bool) -> ExtractMonad<'element, 'element []>
        }

    let makeConsumeModule () : ConsumeModule<'element> = 

        let endOfInput : ExtractMonad<unit, 'element []> = 
            peek (fun pos arr -> pos > arr.Length) >>= fun ans ->
            if ans then 
                mreturn ()
            else extractError "end of input"
            

        let getItem : ExtractMonad<'element, 'element []> = 
            consume1 (fun ix arr -> arr.[ix])

        let getItems (count:int) : ExtractMonad<'element [], 'element []> = 
            consume1 (fun ix arr -> arr.[ix .. ix+count])
         
        /// Doesn't increase the cursor position.
        let getInput : ExtractMonad<'element [], 'element []> = 
            peek (fun ix arr -> arr.[ix ..])

        let inputCount : ExtractMonad<int, 'element []> =  
            peek (fun ix arr -> arr.Length - ix)

        let satisfy (test:'element -> bool) : ExtractMonad<'element, 'element []> = 
            getItem >>= fun ans ->
            if test ans then 
                mreturn ans 
            else 
                extractError "satisfy"

        { 
            EndOfInput = endOfInput
            GetItem = getItem
            GetItems = getItems
            Position = getPosition ()
            GetInput = getInput
            InputCount = inputCount
            Satisfy = satisfy
        }



    //
