module Tables.Tests

open Fable.Mocha
open Fable.ExcelJs

[<Literal>]
let private ws_name = "MySheet1"

let private v_arr = [| Some <| box "TestEntry"; Some <| box "Test"; Some <| box 12; Some "Anything" |]

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
    ]

