module Xlsx.Tests

open Fable.Core.JsInterop
open Fable.Mocha
open Fable.Core
open Fable.Core.JS
open Fable.Core.JsInterop
open Fable.ExcelJs

module TestWorkbooks =
    
    [<Literal>]
    let private ws_name = "MySheet1"

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
        let table_t = Table("MyTable","A1",cols,rows)
        let table = ws.addTable(table_t)
        wb

/// https://github.com/Zaid-Ajaj/Fable.Mocha/issues/69
let main = testList "Xlsx" [ ]

