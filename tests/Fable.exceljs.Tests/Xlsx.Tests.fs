module Xlsx.Tests

open Fable.Mocha
open Fable.Core
open Fable.ExcelJs

module TestWorkbooks =
    
    let [<Literal>] ws_name = "MySheet1"

    let [<Literal>] table_name = "MyTable"

    let Workbook1 =
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        let cols = [|
            TableColumn("Column 1 Test")
            TableColumn("Column 2 Test")
            TableColumn("Column 3 Test")
        |]
        let rows =
            [|
                for i in 0 .. 3 do
                    yield
                        [|box $"Row {i}"; box i; (i%2 |> fun x -> x = 1 |> box)|]
            |]
        let table_t = Table(table_name,"A1",cols,rows)
        let table = ws.addTable(table_t)
        wb

let [<Literal>] File1Path = @"C:\Users\Kevin\source\repos\Fable.Exceljs\tests\Fable.Exceljs.JsNativeTests/ReadWriteFableTest.xlsx"

let tests_write = testList "write" [
    testAsync "ensure async" {
        do! Async.Sleep 300
        Expect.isTrue true ""
    }
    testAsync "write file" {
        do! TestWorkbooks.Workbook1.xlsx.writeFile File1Path |> Async.AwaitPromise
        Expect.isTrue true ""
    }  
]

let tests_read = testList "read" [
    testAsync "read file" {
        let wb = ExcelJs.Excel.Workbook()
        do! wb.xlsx.readFile File1Path |> Async.AwaitPromise
        let ws = wb.getWorksheet(TestWorkbooks.ws_name)
        Expect.equal ws.name TestWorkbooks.ws_name "ws.name"
        let table = ws.getTable(TestWorkbooks.table_name)
        Expect.equal table.name TestWorkbooks.table_name "table.name"
    }  
]

/// https://github.com/Zaid-Ajaj/Fable.Mocha/issues/69
let main = testSequenced <| testList "Xlsx" [ 
    tests_write
    tests_read
]

