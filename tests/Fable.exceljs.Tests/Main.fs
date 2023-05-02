module Fable.exceljs.Tests

open Fable.Mocha

let Main = testList "Main" [
    Tables.Tests.main
]

let [<EntryPoint>] main argv = Mocha.runTests Main
