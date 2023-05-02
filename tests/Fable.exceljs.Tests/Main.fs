module Fable.exceljs.Tests

#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto

[<Tests>]
#endif
let Main = testList "Main" [
    Tests.tests
]

let [<EntryPoint>] main argv = 
    #if FABLE_COMPILER
    Mocha.runTests Main
    #else
    Tests.runTestsWithCLIArgs [] argv Main
    #endif
