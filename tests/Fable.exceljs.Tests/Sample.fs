module Tests

open Fable.exceljs

#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto

[<Tests>]
#endif
let tests =
  testList "samples" [
    testCase "universe exists (╭ರᴥ•́)" <| fun _ ->
      Expect.equal Fable.exceljs.Say.x 0 ""
  ]
