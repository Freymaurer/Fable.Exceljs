module Row.Tests

open Fable.Mocha
open Fable.ExcelJs
open Fable.Core.JsInterop

[<Literal>]
let private ws_name = "MySheet1"

let v_arr(c:char) = [| for i in 1 .. 10 do yield Some <| box $"{c}{i}" |]
let testArr = [|for i in 1 .. 10 do yield Some <| box "CHANGED"|]
let row2_values = v_arr 'a'
let row3_values = v_arr 'b'
let row4_values = v_arr 'c'
let row5_values = v_arr 'd'

let main = testList "Row" [
    testCase "addRow obj option []" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let v_arr = [| None; Some <| box "Test"; Some <| box 12; Some "Anything" |]
        let row = ws.addRow(v_arr)
        Expect.equal row.values.[1..] v_arr ""
    testCase "addRow obj" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let r = ws.addRow({|id1 = "Test"; id2 = 2; id3 = 4; id4 = None; id5 = None; id7 = "Test me too"|})
        let r2 = ws.addRow(createObj [
            "id1" ==> "Test"
            "id2" ==> 2
            "id3" ==> 4
            "id4" ==> None
            "id5" ==> None
            "id7" ==> "Test me too"
        ])
        Expect.equal r.values r2.values "Check difference between anom record type and createObj"
        Expect.equal r.values.[1] (Some "Test") "1"
        Expect.equal r.values.[2] (Some 2) "2"
        Expect.equal r.values.[3] (Some 4) "3"
        Expect.equal r.values.[4] (None) "4"
        Expect.equal r.values.[5] (None) "5"
        Expect.equal r.values.[6] (None) "6"
        Expect.equal r.values.[7] (Some "Test me too") "7"
    testCase "addRows obj" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let row_values = [|
            box {|id1 = "Test"; id2 = 2; id3 = 4; id4 = None; id5 = None; id7 = "Test me too"|};
            createObj [
                "id1" ==> "Test"
                "id2" ==> 2
                "id3" ==> 4
                "id4" ==> None
                "id5" ==> None
                "id7" ==> "Test me too"
            ]
        |]
        let rows = ws.addRows(row_values)
        /// 1 bases index, and 1 is header not new rows
        let rows_get = [|ws.getRow 2; ws.getRow 3|]
        let r = rows_get.[0]
        Expect.equal rows rows_get "Compare set get"
        Expect.equal r.values.[1] (Some "Test") "1"
        Expect.equal r.values.[2] (Some 2) "2"
        Expect.equal r.values.[3] (Some 4) "3"
        Expect.equal r.values.[4] (None) "4"
        Expect.equal r.values.[5] (None) "5"
        Expect.equal r.values.[6] (None) "6"
        Expect.equal r.values.[7] (Some "Test me too") "7"
    testCase "eachRow" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let row2 = ws.addRow row2_values
        let row3 = ws.addRow row3_values
        let row4 = ws.addRow row4_values
        let row5 = ws.addRow row5_values
        ws.eachRow(fun (r,i) ->
            r.values <- testArr
        )
        Expect.equal (ws.getRow(2).values.[1..]) testArr "row1"
        Expect.equal (ws.getRow(3).values.[1..]) testArr "row2"
        Expect.equal (ws.getRow(4).values.[1..]) testArr "row3"
        Expect.equal (ws.getRow(5).values.[1..]) testArr "row4"
    testCase "insertRow" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let row2 = ws.addRow row2_values
        let row3 = ws.addRow row3_values
        let row4 = ws.addRow row4_values
        let row5 = ws.addRow row5_values
        let insertedRow = ws.insertRow(2, testArr)
        Expect.equal (ws.getRow(2).values) insertedRow.values "isInserted"
        Expect.equal (ws.getRow(3).values.[1..]) row2_values "isShifted"
        Expect.equal (ws.getRow(4).values.[1..]) row3_values "isShifted2"
    testCase "insertRows" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let row2 = ws.addRow row2_values
        let row3 = ws.addRow row3_values
        let row4 = ws.addRow row4_values
        let row5 = ws.addRow row5_values
        let insertedRow = ws.insertRows(2, [|testArr; testArr|])
        Expect.equal (ws.getRow(2).values.[1..]) testArr "isInserted"
        Expect.equal (ws.getRow(3).values.[1..]) testArr "isInserted2"
        Expect.equal (ws.getRow(4).values.[1..]) row2_values "isShifted"
    testCase "spliceRows" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let row2 = ws.addRow row2_values
        let row3 = ws.addRow row3_values
        let row4 = ws.addRow row4_values
        let row5 = ws.addRow row5_values
        Expect.hasLength ws.rows 5 "rowLength before splice"
        ws.spliceRows (3,1)
        Expect.hasLength (ws.rows |> Array.filter (isNull >> not)) 4 "rowLength after splice"
        Expect.equal (ws.getRow(2).values.[1..]) row2_values "above splice"
        Expect.equal (ws.getRow(3).values.[1..]) row4_values "below splice"
    testCase "spliceRows with insert" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let row2 = ws.addRow row2_values
        let row3 = ws.addRow row3_values
        let row4 = ws.addRow row4_values
        let row5 = ws.addRow row5_values
        Expect.hasLength ws.rows 5 "rowLength before splice"
        ws.spliceRows (3,1,[|testArr; testArr|])
        Expect.hasLength (ws.rows |> Array.filter (isNull >> not)) 6 "rowLength after splice"
        Expect.equal (ws.getRow(2).values.[1..]) row2_values "above splice"
        Expect.equal (ws.getRow(3).values.[1..]) testArr "below splice insert 1"
        Expect.equal (ws.getRow(4).values.[1..]) testArr "below splice insert 2"
        Expect.equal (ws.getRow(5).values.[1..]) row4_values "below splice"
    testCase "duplicateRow single" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let v_arr(c:char) = [| for i in 1 .. 10 do yield Some <| box $"{c}{i}" |]
        let testArr = [|for i in 1 .. 10 do yield Some <| box "CHANGED"|]
        let row2_values = v_arr 'a'
        let row3_values = v_arr 'b'
        let row4_values = v_arr 'c'
        let row5_values = v_arr 'd'
        let row2 = ws.addRow row2_values
        let row3 = ws.addRow row3_values
        let row4 = ws.addRow row4_values
        let row5 = ws.addRow row5_values
        Expect.hasLength ws.rows 5 "number of Rows"
        ws.duplicateRow(2)
        Expect.equal (ws.getRow(2).values) (ws.getRow(3).values) "duplicate rows are equal"
        Expect.equal (ws.getRow(2).values.[1..]) row2_values "row2"
        Expect.equal (ws.getRow(3).values.[1..]) row2_values "row3"
        Expect.equal (ws.getRow(4).values.[1..]) row3_values "row4"
        Expect.equal (ws.getRow(5).values.[1..]) row4_values "row5"
        Expect.equal (ws.getRow(6).values.[1..]) row5_values "row6"
        Expect.hasLength ws.rows 6 "number of Rows after"
    testCase "duplicateRow single replace" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let row2 = ws.addRow row2_values
        let row3 = ws.addRow row3_values
        let row4 = ws.addRow row4_values
        let row5 = ws.addRow row5_values
        Expect.hasLength ws.rows 5 "number of Rows"
        ws.duplicateRow(2, false)
        Expect.equal (ws.getRow(2).values) (ws.getRow(3).values) "duplicate rows are equal"
        Expect.equal (ws.getRow(2).values.[1..]) row2_values "row2"
        Expect.equal (ws.getRow(3).values.[1..]) row2_values "row3"
        Expect.equal (ws.getRow(4).values.[1..]) row4_values "row4"
        Expect.equal (ws.getRow(5).values.[1..]) row5_values "row5"
        Expect.hasLength ws.rows 5 "number of Rows"
    testCase "duplicateRow multiple" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let row2 = ws.addRow row2_values
        let row3 = ws.addRow row3_values
        let row4 = ws.addRow row4_values
        let row5 = ws.addRow row5_values
        Expect.hasLength ws.rows 5 "number of Rows"
        ws.duplicateRow(2, 3)
        Expect.equal (ws.getRow(2).values) (ws.getRow(3).values) "duplicate rows are equal"
        Expect.equal (ws.getRow(2).values.[1..]) row2_values "row2"
        Expect.equal (ws.getRow(3).values.[1..]) row2_values "row3"
        Expect.equal (ws.getRow(4).values.[1..]) row2_values "row4"
        Expect.equal (ws.getRow(5).values.[1..]) row2_values "row5"
        Expect.equal (ws.getRow(6).values.[1..]) row3_values "row6"
        Expect.hasLength ws.rows 8 "number of Rows"
    testCase "duplicateRow multiple replace" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let v_arr(c:char) = [| for i in 1 .. 10 do yield Some <| box $"{c}{i}" |]
        let testArr = [|for i in 1 .. 10 do yield Some <| box "CHANGED"|]
        let row2_values = v_arr 'a'
        let row3_values = v_arr 'b'
        let row4_values = v_arr 'c'
        let row5_values = v_arr 'd'
        let row2 = ws.addRow row2_values
        let row3 = ws.addRow row3_values
        let row4 = ws.addRow row4_values
        let row5 = ws.addRow row5_values
        Expect.hasLength ws.rows 5 "number of Rows"
        ws.duplicateRow(2,3,false)
        Expect.equal (ws.getRow(2).values) (ws.getRow(3).values) "duplicate rows are equal"
        Expect.equal (ws.getRow(2).values.[1..]) row2_values "row2"
        Expect.equal (ws.getRow(3).values.[1..]) row2_values "row3"
        Expect.equal (ws.getRow(4).values.[1..]) row2_values "row4"
        Expect.equal (ws.getRow(5).values.[1..]) row2_values "row5"
        Expect.hasLength ws.rows 5 "number of Rows"
    testCase "duplicateRow multiple replace more than previous existing" <| fun _ ->
        let wb = ExcelJs.exceljs.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let v_arr(c:char) = [| for i in 1 .. 10 do yield Some <| box $"{c}{i}" |]
        let testArr = [|for i in 1 .. 10 do yield Some <| box "CHANGED"|]
        let row2_values = v_arr 'a'
        let row3_values = v_arr 'b'
        let row4_values = v_arr 'c'
        let row5_values = v_arr 'd'
        let row2 = ws.addRow row2_values
        let row3 = ws.addRow row3_values
        let row4 = ws.addRow row4_values
        let row5 = ws.addRow row5_values
        Expect.hasLength ws.rows 5 "number of Rows"
        ws.duplicateRow(2,4,false)
        Expect.equal (ws.getRow(2).values) (ws.getRow(3).values) "duplicate rows are equal"
        Expect.equal (ws.getRow(2).values.[1..]) row2_values "row2"
        Expect.equal (ws.getRow(3).values.[1..]) row2_values "row3"
        Expect.equal (ws.getRow(4).values.[1..]) row2_values "row4"
        Expect.equal (ws.getRow(5).values.[1..]) row2_values "row5"
        Expect.equal (ws.getRow(6).values.[1..]) row2_values "row6"
        Expect.hasLength ws.rows 6 "number of Rows"
    ]