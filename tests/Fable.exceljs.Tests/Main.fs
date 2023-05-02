module Fable.exceljs.Tests

open Fable.Mocha

let Main = testList "Main" [
    Tables.Tests.main
    Cell.Tests.main
    Row.Tests.main
    Column.Tests.main
    Worksheet.Tests.main
    Workbook.Tests.main
]

let [<EntryPoint>] main argv = Mocha.runTests Main
