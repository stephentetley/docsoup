# Monad notes

FParsec places the answer type as the first argument in the type constructor
and user state as the second:

~~~
Parser<'a,'u>
~~~

Haskell always favours placing the answer type as the last argument to work
with the Functor etc. classes. Parsec also has an extra parameter for the
input stream type:

~~~
GenParser tok st a
~~~

## Standard operations

+---------------+---------------+--------------------+
| FSharp        | Haskell       | FParsec            |
+:==============+:==============+:===================+
| `return`      | `return`      | `preturn`          |
+---------------+---------------+--------------------+
| &nbsp;        | `>>=`         | `>>=`              |
+---------------+---------------+--------------------+
| `mzero`       | &nbsp;        | `pzero`            |
+---------------+---------------+--------------------+
