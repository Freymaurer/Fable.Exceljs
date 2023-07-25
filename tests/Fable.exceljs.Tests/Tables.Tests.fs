module Tables.Tests

open Fable.Mocha
open Fable.ExcelJs

open Fable.Core.JsInterop
open Fable.Core

[<Literal>]
let private ws_name = "MySheet1"

let private v_arr = [| Some <| box "TestEntry"; Some <| box "Test"; Some <| box 12; Some "Anything" |]

[<Emit("console.log($0)")>]
let log(obj:obj) = jsNative

let main = testList "Tables" [
    testCase "Table()" <| fun _ ->
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
        Expect.pass ()
    testCase "tableRef" <| fun _ ->
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
        Expect.equal table.ref "A1" "ref"
        Expect.equal (table.tableRef) "A1:C5" "tableRef"
    testCase "tableRef, not equal" <| fun _ ->
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
        Expect.equal table.ref "A1" "ref"
        Expect.notEqual (table.tableRef) "A1:C6" "tableRef"
    testCase "ws.addTable" <| fun _ ->
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
        Expect.equal table.name "MyTable" table.name
    testCase "ws.addTable" <| fun _ ->
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
        let getTable = ws.getTable("MyTable")
        Expect.equal getTable table "getTable"
    testCase "getTables" <| fun _ ->
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
        let getTables = ws.getTables()
        Expect.equal getTables.Length 1 "count"
        Expect.equal getTables.[0] table "equal"
    testCase "removeTables" <| fun _ ->
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
        let getTables = ws.getTables()
        Expect.equal getTables.Length 1 "count"
        ws.removeTable("MyTable")
        let getTables2 = ws.getTables()
        Expect.equal getTables2.Length 0 "count"
    ]

