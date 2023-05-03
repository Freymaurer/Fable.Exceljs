module Xlsx.Tests

open Fable.Core.JsInterop
open Fable.Mocha
open Fable.Core
open Fable.Core.JS
open Fable.Core.JsInterop
open Fable.ExcelJs

[<Literal>]
let MinimalPath = @"C:\Users\Kevin\source\repos\Fable.exceljs\tests\files\MinimalTest.xlsx"

let main = testList "Xlsx" [
    testAsync "async test" {
        do! Async.Sleep 300
        Expect.isTrue true ""
    }
    testCaseAsync "two" <| async {
        do! Async.Sleep 1000
        Expect.isTrue true ""
    }
    testAsync "Read Xlsx" {
        let workbook = ExcelJs.Excel.Workbook()
        Expect.passWithMsg "Create workbook"
        do! workbook.xlsx.readFile(MinimalPath)
        Expect.passWithMsg "Read File"
        let worksheet = workbook.getWorksheet("sheet1");
        Expect.passWithMsg "Get Worksheet"
    }
]

