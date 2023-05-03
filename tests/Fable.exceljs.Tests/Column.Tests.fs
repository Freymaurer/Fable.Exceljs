module Column.Tests

open Fable.Mocha
open Fable.ExcelJs


[<Literal>]
let ws_name = "MySheet1"

let main = testList "Column" [
    testCase "columns update" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        Expect.hasLength ws.columns 10 "hasLength"
        Expect.equal ws.columnCount 10 "columnCount"
    testCase "column props" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let col = ws.getColumn("C")
        let hidden = true
        col.hidden <- hidden
        Expect.equal col.hidden hidden "hidden"
        Expect.equal col.letter "C" "letter"
        Expect.equal col.header "Header3" "header"
    testCase "getColumn" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            for i in 1 .. 10 do
                yield Column.create($"Header{i}", $"id{i}", 10, 11, false)
        |]
        let col2_1 = ws.getColumn("B")
        let col2_2 = ws.getColumn(2)
        let col2_3 = ws.getColumn("id2")
        Expect.equal col2_1 col2_2 "equal 1"
        Expect.equal col2_1 col2_2 "equal 1"
        Expect.equal col2_1 col2_3 "equal 2"
    testCase "eachCell" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        let col = ws.getColumn(1)
        col.values <- [|Some "Row1"; Some "Row2"; Some "Row3"|]
        col.eachCell(fun (c,rowi) ->
            c.value <- Some rowi
            ()
        )
        let nonEmptyValues = col.values |> Seq.choose id |> Array.ofSeq
        Expect.hasLength nonEmptyValues 3 "hasLength"
        Expect.equal nonEmptyValues.[0] 1 "1"
        Expect.equal nonEmptyValues.[1] 2 "2"
        Expect.equal nonEmptyValues.[2] 3 "3"
    testCase "eachCell_IncludeEmpty" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        let col = ws.getColumn(1)
        col.values <- [|Some "Row1"; None; Some "Row3"|]
        col.eachCell(true, fun (c,rowi) ->
            c.value <- Some rowi
            ()
        )
        let skipZeroBasedIndex = col.values |> Seq.tail |> Array.ofSeq
        Expect.hasLength skipZeroBasedIndex 3 "hasLength"
        Expect.equal skipZeroBasedIndex.[0] (Some 1) "1"
        Expect.equal skipZeroBasedIndex.[1] (Some 2) "2"
        Expect.equal skipZeroBasedIndex.[2] (Some 3) "3"
    testCase "eachCell_IncludeEmpty_Not" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        let col = ws.getColumn(1)
        col.values <- [|Some "Row1"; None; Some "Row3"|]
        col.eachCell(false, fun (c,rowi) ->
            c.value <- Some rowi
            ()
        )
        let skipZeroBasedIndex = col.values |> Seq.tail |> Array.ofSeq
        Expect.hasLength skipZeroBasedIndex 3 "hasLength"
        Expect.equal skipZeroBasedIndex.[0] (Some 1) "1"
        Expect.equal skipZeroBasedIndex.[1] (None) "None"
        Expect.equal skipZeroBasedIndex.[2] (Some 3) "3"
]