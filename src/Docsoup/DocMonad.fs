// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace Docsoup


module DocMonad = 
    
    open Microsoft.Office.Interop

    type ErrMsg = string

    type Answer<'a, 'state> = Result<'a * 'state, ErrMsg>

    type DocMonad<'handle, 'cursor, 'a> = 
        DocMonad of ('handle -> 'cursor -> Answer<'a, 'cursor>)
        


    let inline private apply1 (ma: DocMonad<'handle, 'cursor, 'a>) 
                              (handle: 'handle)
                              (position: 'cursor) : Answer<'a, 'cursor>= 
        let (DocMonad f) = ma in f handle position

        
    let inline mreturn (x:'a) : DocMonad<'handle, 'cursor, 'a> = 
        DocMonad <| fun _ pos -> Ok (x, pos)

    let inline private bindM (ma:DocMonad<'handle, 'cursor, 'a>) 
        (f :'a -> DocMonad<'handle, 'cursor, 'b>) : DocMonad<'handle, 'cursor, 'b> =
            DocMonad <| fun handle pos -> 
                match apply1 ma handle pos with
                | Error msg -> Error msg
                | Ok (a, pos1) -> apply1 (f a) handle pos1

    let inline private zeroM () : DocMonad<'handle, 'cursor,'a> = 
        DocMonad <| fun _ _ -> Error "zeroM"

    /// "First success"
    let inline private combineM (ma:DocMonad<'handle, 'cursor,'a>) 
                                (mb:DocMonad<'handle, 'cursor,'a>) : DocMonad<'handle, 'cursor,'a> = 
        DocMonad <| fun handle pos -> 
            match apply1 ma handle pos with
            | Error msg -> apply1 mb handle pos
            | Ok (a,pos1) -> Ok (a,pos1)

    let inline private delayM (fn:unit -> DocMonad<'handle, 'cursor,'a>) : DocMonad<'handle, 'cursor,'a> = 
        bindM (mreturn ()) fn 

    type DocMonadBuilder<'handle, 'cursor>() = 
        member self.Return(x:'a) : DocMonad<'handle, 'cursor, 'a>  = mreturn x
        member self.Bind (p:DocMonad<'handle,'cursor, 'a> , f: 'a -> DocMonad<'handle,'cursor, 'b>) : DocMonad<'handle, 'cursor, 'b>         = bindM p f
        member self.Zero () : DocMonad<'handle,'cursor, 'a> = zeroM ()
        member self.Combine (ma: DocMonad<'handle,'cursor, 'a>, mb: DocMonad<'handle,'cursor, 'a>) : DocMonad<'handle,'cursor, 'a>  = combineM ma mb
        member self.Delay (fn:unit -> DocMonad<'handle, 'cursor, 'a>) : DocMonad<'handle, 'cursor,'a> = delayM fn
        member self.ReturnFrom(ma:DocMonad<'handle, 'cursor,'a>) : DocMonad<'handle, 'cursor, 'a> = ma

    type DocParserBuilder = DocMonadBuilder<Word.Document, int> 

    let (docParser:DocParserBuilder) = new DocMonadBuilder<Word.Document, int>()

    type DocParser<'a> = DocMonad<Word.Document, int, 'a> 
    
    let action1 : DocParser<unit> = 
        docParser { 
            return ()
        }

    type TablesParserBuilder = DocMonadBuilder<Word.Table [], int> 

    let (tablesParser:TablesParserBuilder) = new DocMonadBuilder<Word.Table [], int>()

    type TablesParser<'a> = DocMonad<Word.Table [], int, 'a> 

    let action2 : TablesParser<unit> = 
        tablesParser { 
            return ()
        }