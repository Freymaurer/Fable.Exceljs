module Worksheet.Tests

open Fable.Mocha
open Fable.ExcelJs

[<Literal>]
let ws_name = "MySheet1"

let main = testList "Worksheet" [
    testCase "Worksheet properties" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        let name = ws_name + "TEST"
        let state = WorksheetState.Hidden
        let properties = WorksheetProperties.create(
            outlineLevelCol = 12,
            outlineLevelRow = 13,
            defaultRowHeight = 14,
            defaultColWidth = 15,
            dyDescent = 16
        )
        ws.name <- name
        ws.state <- state
        ws.properties <- properties
        Expect.equal ws.name name "name"
        Expect.equal ws.state state "state"
        Expect.equal ws.properties properties "properties"
        Expect.equal ws.lastRowNumber 0 "lastRowNumer"
        Expect.isNone ws.lastRow "lastRow"
        Expect.isNone ws.lastColumn "lastColumn"
    testCase "Last row isSome" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            Column.create("HeaderUp", "id")
            Column.create("HeaderDown", "id2")
        |]
        Expect.isSome ws.lastRow "lastRow isSome"
        Expect.equal ws.lastRowNumber 1 "lastRow isSome"
        Expect.equal ws.lastRow.Value.cellCount 2 "lastRow.Value.values"
    testCase "Last column isSome" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            Column.create("HeaderUp", "id")
            Column.create("HeaderDown", "id2")
        |]
        Expect.isSome ws.lastColumn "lastRow isSome"
    testCase "Worksheet columns" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            Column.create("HeaderUp", "id")
            Column.create("HeaderDown", "id2")
        |]
        Expect.hasLength ws.columns 2 "column length"
    testCase "Worksheet row/column count" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            Column.create("HeaderUp", "id")
            Column.create("HeaderDown", "id2")
        |]
        Expect.equal ws.actualColumnCount 2 "actualColumnCount"
        Expect.equal ws.columnCount 2 "actualColumnCount"
        Expect.equal ws.actualRowCount 1 "actualRowCount"
        Expect.equal ws.rowCount 1 "actualRowCount"
    testCase "Worksheet row/column count 2" <| fun _ ->
        let wb = ExcelJs.Excel.Workbook()
        let ws = wb.addWorksheet(ws_name)
        ws.columns <- [|
            Column.create("HeaderUp", "id")
            Column.create("HeaderDown", "id2")
        |]
        let c = ws.getColumn("B")
        ws.getColumn("C").values <- [|None|]
        c.values <- [|None; None; Some (box "Testme")|]
        Expect.equal ws.actualColumnCount 2 "actualColumnCount"
        Expect.equal ws.columnCount 3 "thirs is empty; columnCount"
        Expect.equal ws.actualRowCount 2 "first and third, second is empty; actualRowCount"
        Expect.equal ws.rowCount 3 " rowCount"
]